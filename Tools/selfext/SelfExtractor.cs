using System;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Collections.Generic;
using System.Windows.Forms;

/// <summary>
/// SmartPilot for Office - GUI installer (WinForms, no console window).
/// The portable ZIP is embedded as the "selfext.zip" manifest resource.
/// Flow: Start -> extract to %LOCALAPPDATA%\Programs\ExcelAIVSTOAddIn
///       (progress bar) -> run register.ps1 via absolute-path PowerShell
///       (output shown in log) -> done page. No admin required.
/// </summary>
class InstallerForm : Form
{
    private Label _status;
    private ProgressBar _bar;
    private TextBox _log;
    private Button _btn;
    private bool _done;

    public InstallerForm()
    {
        Text = "SmartPilot for Office 安装程序";
        Font = new Font("Microsoft YaHei UI", 9F);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(560, 386);

        _status = new Label();
        _status.Text = "点击「开始安装」，加载项将安装到当前用户目录（无需管理员）。";
        _status.Location = new Point(16, 14);
        _status.Size = new Size(528, 22);
        Controls.Add(_status);

        _bar = new ProgressBar();
        _bar.Location = new Point(16, 42);
        _bar.Size = new Size(528, 20);
        Controls.Add(_bar);

        _log = new TextBox();
        _log.Multiline = true;
        _log.ReadOnly = true;
        _log.ScrollBars = ScrollBars.Vertical;
        _log.Font = new Font("Consolas", 8.5F);
        _log.Location = new Point(16, 72);
        _log.Size = new Size(528, 252);
        Controls.Add(_log);

        _btn = new Button();
        _btn.Text = "开始安装";
        _btn.Location = new Point(16, 336);
        _btn.Size = new Size(528, 34);
        _btn.Click += new EventHandler(OnStart);
        Controls.Add(_btn);
    }

    private void OnStart(object sender, EventArgs e)
    {
        if (_done) { Close(); return; }
        _btn.Enabled = false;
        _btn.Text = "正在安装...";
        Thread worker = new Thread(new ThreadStart(RunInstall));
        worker.IsBackground = true;
        worker.Start();
    }

    private void SetStatus(string s) { Safe(delegate { _status.Text = s; }); }
    private void AppendLog(string s) { Safe(delegate { _log.AppendText(s + "\r\n"); }); }

    private void Safe(Action a)
    {
        try { BeginInvoke(a); } catch (ObjectDisposedException) { }
    }

    private void Finish(bool ok, string msg)
    {
        Safe(delegate { _bar.Value = _bar.Maximum; });
        Safe(delegate { _status.Text = msg; _status.ForeColor = ok ? Color.Green : Color.Red; });
        Safe(delegate
        {
            _done = true;
            _btn.Enabled = true;
            _btn.Text = ok ? "完成" : "关闭";
        });
    }

    private static void CreateShortcut(string lnkPath, string targetPath, string workDir, string description)
    {
        Type shellType = Type.GetTypeFromProgID("WScript.Shell");
        object shell = Activator.CreateInstance(shellType);
        object lnk = shellType.InvokeMember("CreateShortcut",
            System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { lnkPath });
        Type lnkType = lnk.GetType();
        lnkType.InvokeMember("TargetPath", System.Reflection.BindingFlags.SetProperty,
            null, lnk, new object[] { targetPath });
        lnkType.InvokeMember("WorkingDirectory", System.Reflection.BindingFlags.SetProperty,
            null, lnk, new object[] { workDir });
        lnkType.InvokeMember("Description", System.Reflection.BindingFlags.SetProperty,
            null, lnk, new object[] { description });
        lnkType.InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod, null, lnk, null);
    }

    private void CreateStartMenu(string installDir)
    {
        try
        {
            string smDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.StartMenu),
                "Programs", "SmartPilot for Office");
            Directory.CreateDirectory(smDir);
            CreateShortcut(Path.Combine(smDir, "卸载 SmartPilot（Office 插件）.lnk"),
                Path.Combine(installDir, "uninstall.bat"), installDir, "卸载 SmartPilot Office 插件");
            AppendLog("已创建开始菜单卸载入口（开始菜单 -> SmartPilot for Office）。");
        }
        catch (Exception ex)
        {
            AppendLog("创建开始菜单卸载入口失败（不影响使用）: " + ex.Message);
        }
    }

    private static readonly string[] OfficeNames = new string[] { "excel", "winword", "powerpnt", "OcrHost" };

    private static string OfficeDisplayName(string n)
    {
        if (n == "excel") return "Excel";
        if (n == "winword") return "Word";
        if (n == "OcrHost") return "OcrHost（SmartPilot OCR 引擎）";
        return "PowerPoint";
    }

    private static string DetectOffice()
    {
        List<string> found = new List<string>();
        foreach (string n in OfficeNames)
        {
            if (Process.GetProcessesByName(n).Length > 0) found.Add(OfficeDisplayName(n));
        }
        return found.Count > 0 ? String.Join("、", found.ToArray()) : null;
    }

    /// <summary>Asks the user to close running Office apps (optionally closing them).</summary>
    private bool EnsureOfficeClosed()
    {
        while (true)
        {
            string running = DetectOffice();
            if (running == null) return true;
            DialogResult dr = MessageBox.Show(
                "检测到以下程序正在运行：\n\n" + running +
                "\n\n继续安装需要关闭它们（未保存的修改将丢失）。\n\n是否自动关闭并继续？",
                "SmartPilot for Office 安装程序",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
            if (dr == DialogResult.No) return false;
            foreach (string n in OfficeNames)
            {
                foreach (Process p in Process.GetProcessesByName(n))
                {
                    try
                    {
                        p.CloseMainWindow();
                        if (!p.WaitForExit(5000)) p.Kill();
                    }
                    catch (Exception) { }
                }
            }
            Thread.Sleep(1000);
        }
    }

    private void RunInstall()
    {
        try
        {
            if (!EnsureOfficeClosed())
            {
                Finish(false, "安装已取消。请关闭所有 Office 程序后重新运行安装程序。");
                return;
            }
            string baseDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Programs", "ExcelAIVSTOAddIn");
            AppendLog("安装目录: " + baseDir);

            SetStatus("正在解压文件...");
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream stream = asm.GetManifestResourceStream("selfext.zip"))
            {
                if (stream == null) throw new Exception("安装包数据缺失（selfext.zip 资源不存在）");
                using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read))
                {
                    int total = archive.Entries.Count;
                    Safe(delegate { _bar.Maximum = total; });
                    bool extracted = false;
                    while (!extracted)
                    {
                        int done = 0;
                        try
                        {
                            foreach (ZipArchiveEntry entry in archive.Entries)
                            {
                                string fullName = entry.FullName.Replace('/', '\\');
                                if (fullName.Length > 0)
                                {
                                    string dest = Path.Combine(baseDir, fullName);
                                    if (fullName.EndsWith("\\"))
                                    {
                                        Directory.CreateDirectory(dest);
                                    }
                                    else
                                    {
                                        string parent = Path.GetDirectoryName(dest);
                                        if (!String.IsNullOrEmpty(parent)) Directory.CreateDirectory(parent);
                                        entry.ExtractToFile(dest, true);
                                    }
                                }
                                done++;
                                if (done % 20 == 0 || done == total)
                                {
                                    int d = done;
                                    Safe(delegate { _bar.Value = d; });
                                }
                            }
                            extracted = true;
                        }
                        catch (IOException ioEx)
                        {
                            AppendLog("错误: " + ioEx.Message);
                            DialogResult dr = MessageBox.Show(
                                "文件被其他程序占用：\n" + ioEx.Message +
                                "\n\n请关闭 Excel / Word / PowerPoint 及 OCR 引擎（OcrHost）后，点「是」重试，「否」取消安装。",
                                "SmartPilot for Office 安装程序",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                            if (dr == DialogResult.No)
                            {
                                Finish(false, "安装已取消。请关闭所有相关程序后重新运行安装程序。");
                                return;
                            }
                            EnsureOfficeClosed();
                        }
                    }
                }
            }
            AppendLog("解压完成。");

            SetStatus("正在注册 Office 加载项（Excel / Word / PowerPoint）...");
            AppendLog("");
            AppendLog("------ 注册输出 ------");

            string sysDir = Environment.GetFolderPath(Environment.SpecialFolder.System);
            string ps = Path.Combine(sysDir, "WindowsPowerShell", "v1.0", "powershell.exe");
            if (!File.Exists(ps)) ps = "powershell.exe";
            string script = Path.Combine(baseDir, "register.ps1");

            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = ps;
            psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + script + "\"";
            psi.UseShellExecute = false;
            psi.CreateNoWindow = true;
            psi.RedirectStandardOutput = true;
            psi.RedirectStandardError = true;

            bool ok = false;
            using (Process p = Process.Start(psi))
            {
                p.OutputDataReceived += new DataReceivedEventHandler(
                    delegate(object o, DataReceivedEventArgs ev) { if (ev.Data != null) AppendLog(ev.Data); });
                p.ErrorDataReceived += new DataReceivedEventHandler(
                    delegate(object o, DataReceivedEventArgs ev) { if (ev.Data != null) AppendLog(ev.Data); });
                p.BeginOutputReadLine();
                p.BeginErrorReadLine();
                p.WaitForExit();
                ok = (p.ExitCode == 0);
            }
            AppendLog("----------------------");
            AppendLog("");

            if (ok)
            {
                CreateStartMenu(baseDir);
                Finish(true, "安装完成！请完全退出所有 Office 程序（Excel / Word / PowerPoint），然后重新打开。");
            }
            else
                Finish(false, "安装失败，请把上方日志中的红色错误信息反馈给支持人员。");
        }
        catch (Exception ex)
        {
            AppendLog("错误: " + ex.Message);
            Finish(false, "安装失败：" + ex.Message);
        }
    }
}

class SelfExtractor
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new InstallerForm());
    }
}
