using System;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using Microsoft.Win32;
using System.Windows.Forms;

/// <summary>
/// SmartPilot desktop edition - GUI installer (WinForms, no console).
/// Extracts to %LOCALAPPDATA%\Programs\SmartPilotWeb, creates a desktop
/// shortcut to start.bat, and reports whether Python is available
/// (start.bat auto-installs pywin32/openpyxl on first run).
/// </summary>
class WebInstallerForm : Form
{
    private Label _status;
    private ProgressBar _bar;
    private TextBox _log;
    private Button _btn;
    private bool _done;

    public WebInstallerForm()
    {
        Text = "SmartPilot 桌面版安装程序";
        Font = new Font("Microsoft YaHei UI", 9F);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(560, 386);

        _status = new Label();
        _status.Text = "点击「开始安装」，安装到当前用户目录并创建桌面快捷方式（无需管理员）。";
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

    private static string DetectPython()
    {
        try
        {
            ProcessStartInfo pi = new ProcessStartInfo("cmd.exe", "/c python --version");
            pi.UseShellExecute = false;
            pi.CreateNoWindow = true;
            pi.RedirectStandardOutput = true;
            pi.RedirectStandardError = true;
            using (Process p = Process.Start(pi))
            {
                string outp = p.StandardOutput.ReadToEnd().Trim();
                string errp = p.StandardError.ReadToEnd().Trim();
                p.WaitForExit();
                if (outp.Length > 0) return outp;
                if (errp.ToLower().Contains("python")) return errp;
                return null;
            }
        }
        catch (Exception) { return null; }
    }

    private static void CreateShortcut(string lnkPath, string targetPath, string workDir, string description, string iconPath)
    {
        Type shellType = Type.GetTypeFromProgID("WScript.Shell");
        object shell = Activator.CreateInstance(shellType);
        object lnk = shellType.InvokeMember("CreateShortcut",
            BindingFlags.InvokeMethod, null, shell, new object[] { lnkPath });
        Type lnkType = lnk.GetType();
        lnkType.InvokeMember("TargetPath", BindingFlags.SetProperty,
            null, lnk, new object[] { targetPath });
        lnkType.InvokeMember("WorkingDirectory", BindingFlags.SetProperty,
            null, lnk, new object[] { workDir });
        lnkType.InvokeMember("Description", BindingFlags.SetProperty,
            null, lnk, new object[] { description });
        if (!String.IsNullOrEmpty(iconPath))
        {
            lnkType.InvokeMember("IconLocation", BindingFlags.SetProperty,
                null, lnk, new object[] { iconPath + ",0" });
        }
        lnkType.InvokeMember("Save", BindingFlags.InvokeMethod, null, lnk, null);
    }

    private void CreateShortcuts(string installDir)
    {
        string icon = Path.Combine(installDir, "smartpilot.ico");
        try
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            CreateShortcut(Path.Combine(desktop, "SmartPilot.lnk"),
                Path.Combine(installDir, "SmartPilot.vbs"), installDir,
                "SmartPilot 桌面版（浏览器 + 本地桥）", icon);
            AppendLog("已创建桌面快捷方式: " + Path.Combine(desktop, "SmartPilot.lnk"));
        }
        catch (Exception ex)
        {
            AppendLog("创建桌面快捷方式失败（不影响使用，可手动双击 SmartPilot.vbs）: " + ex.Message);
        }
        try
        {
            string smDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.StartMenu),
                "Programs", "SmartPilot");
            Directory.CreateDirectory(smDir);
            CreateShortcut(Path.Combine(smDir, "SmartPilot（桌面版）.lnk"),
                Path.Combine(installDir, "SmartPilot.vbs"), installDir, "启动 SmartPilot 桌面版", icon);
            CreateShortcut(Path.Combine(smDir, "卸载 SmartPilot（桌面版）.lnk"),
                Path.Combine(installDir, "uninstall-web.bat"), installDir, "卸载 SmartPilot 桌面版", icon);
            AppendLog("已创建开始菜单入口（开始菜单 -> SmartPilot）。");
        }
        catch (Exception ex)
        {
            AppendLog("创建开始菜单入口失败（不影响使用）: " + ex.Message);
        }
    }

    private void WriteUninstallKey(string installDir)
    {
        try
        {
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(
                "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\SmartPilot"))
            {
                k.SetValue("DisplayName", "SmartPilot（桌面版）");
                k.SetValue("DisplayVersion", "1.0.0");
                k.SetValue("Publisher", "SmartPilot");
                k.SetValue("InstallLocation", installDir);
                k.SetValue("UninstallString", Path.Combine(installDir, "uninstall-web.bat"));
                k.SetValue("NoModify", 1, RegistryValueKind.DWord);
                k.SetValue("NoRepair", 1, RegistryValueKind.DWord);
            }
            AppendLog("已注册卸载入口（设置 -> 应用 -> SmartPilot（桌面版））。");
        }
        catch (Exception ex)
        {
            AppendLog("注册卸载入口失败（可用安装目录内 uninstall-web.bat 卸载）: " + ex.Message);
        }
    }

    private void RunInstall()
    {
        try
        {
            string baseDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Programs", "SmartPilotWeb");
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
                                if (done % 10 == 0 || done == total)
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
                                "\n\n请关闭正在运行的 SmartPilot 桌面版（黑色桥窗口）后，点「是」重试，「否」取消安装。",
                                "SmartPilot 桌面版安装程序",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                            if (dr == DialogResult.No)
                            {
                                Finish(false, "安装已取消。请关闭所有相关程序后重新运行安装程序。");
                                return;
                            }
                        }
                    }
                }
            }
            AppendLog("解压完成。");

            SetStatus("正在检测运行环境...");
            string py = DetectPython();
            AppendLog("");
            if (py != null)
            {
                AppendLog("Python 环境: " + py);
                AppendLog("首次运行 start.bat 时会自动安装 pywin32 / openpyxl 依赖（需要联网）。");
            }
            else
            {
                AppendLog("[提示] 未检测到 Python。请到 python.org 安装 Python 3.10+");
                AppendLog("       （安装时务必勾选 \"Add Python to PATH\"），否则桌面版无法启动。");
            }
            AppendLog("");

            CreateShortcuts(baseDir);
            WriteUninstallKey(baseDir);

            Finish(true, "安装完成！双击桌面上的 SmartPilot 图标启动（首次启动会自动安装 Python 依赖）");
        }
        catch (Exception ex)
        {
            AppendLog("错误: " + ex.Message);
            Finish(false, "安装失败：" + ex.Message);
        }
    }
}

class WebSelfExtractor
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new WebInstallerForm());
    }
}
