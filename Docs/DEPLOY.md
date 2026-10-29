# 上线 / 发布指南（ExcelAIVSTOAddIn）

本插件是 VSTO 加载项（ClickOnce 分发），同时服务 **Excel / Word / PowerPoint** 三个宿主，
前端用 WebView2 承载。本文档汇总"上线布局前必须完成"的步骤与高风险陷阱，配合
`Tools/release-check.ps1` 自动闸门使用。

---

## 0. 一键预检

```powershell
cd 仓库根
Tools\release-check.ps1            # 全量（含测试，可作 CI 发布闸门）
Tools\release-check.ps1 -SkipTests # 跳过耗时测试
```

脚本会检查：签名证书是否临时、版本号、各宿主 `bin\Release\UI` 的 `?v=` 戳一致性、
MCP 工具一致性、前端 lint、自动化测试、法务文档、发布目录是否夹带密钥、三宿主版本是否一致、
`InstallUrl` 是否已填。任意 ERROR 返回退出码 1，禁止发布。

日常 CI 使用 `Tools\release-check.ps1 -Ci -SkipTests`：临时证书和空 `InstallUrl` 记为警告，
其余项仍阻断。打 `v*` tag 或在 Actions 里手动跑 **Release gate (strict)** 时不带 `-Ci`，
证书和更新地址未就绪则失败。

---

## 1. 替换正式签名证书（最高优先级）

当前工程使用**临时开发证书**，仅对开发机信任，最终用户安装会被 SmartScreen / Office
拦截或报"未知发布者"。

- 临时证书指纹：`ED1AC12D8BB72993C13AFA541F279849D1A1DF27`
- 临时 key 文件：`ExcelAIVSTOAddIn_TemporaryKey.pfx` / `ExcelAIVSTOAddIn.snk`

替换为正式代码签名证书后，更新 `ExcelAIVSTOAddIn.vbproj` 两处：

```xml
<ManifestKeyFile>你的正式证书.pfx</ManifestKeyFile>
<ManifestCertificateThumbprint>正式证书指纹</ManifestCertificateThumbprint>
```

> 代码签名证书需为受信任 CA 签发（如 DigiCert / Sectigo 的 EV 或 OV 代码签名证书）。
> EV 证书可免 SmartScreen 未知警告；OV 首次发布仍可能短暂弹警告。

---

## 1.1 填写 InstallUrl

把 Excel 与 Word 工程的 `InstallUrl` 设成同一个公网 HTTPS 目录，已安装的客户端靠它找更新。
留空时，不带 `-Ci` 的 `release-check.ps1` 会失败。PowerPoint 当前按 COM 加载项注册，不使用这条地址。

购买证书、填写 `InstallUrl`、把 `Legal\` 里的「【运营方】」换成营业执照主体，这三件事都要人工完成，脚本不能代填。

---

## 2. 递增应用版本号（ClickOnce 更新依赖）

ClickOnce 靠 `ApplicationVersion` + `ApplicationRevision` 判断"是否最新"。

`ExcelAIVSTOAddIn.vbproj`：

```xml
<ApplicationVersion>1.0.0.0</ApplicationVersion>
<ApplicationRevision>0</ApplicationRevision>
```

每次对外发布**必须递增**（建议修订号 +1，大版本变更改 `ApplicationVersion`），
否则已安装用户收不到更新提示。

---

## 3. Release 构建（三宿主分别构建）

三个宿主是独立工程，各自产出并同步 `bin\Release\UI` 与 `Config`：

| 宿主 | 工程 |
|---|---|
| Excel | `ExcelAIVSTOAddIn.vbproj` |
| Word  | `Hosts\WordAIVSTOAddIn\WordAIVSTOAddIn.vbproj` |
| PPT   | `Hosts\PptAIVSTOAddIn\PptAIVSTOAddIn.vbproj` |

构建命令（任一工程）：

```powershell
& 'VS路径\MSBuild\Current\Bin\MSBuild.exe' <工程>.vbproj /p:Configuration=Release /t:Build
```

构建后自动执行（见各 vbproj / `Hosts/Shared/SmartPilot.Build.targets`）：

- `SyncUIFiles`：robocopy `UI/` → `bin\Release\UI`
- `SyncConfigFiles`（Word/PPT）：同步 `Config\*.json`
- `SyncSkillsFiles`（Excel）：同步 `Skills\*.md`
- `SyncBridgeFiles`：部署 `py-excel-bridge.py` + UI 到本机 `LocalAppData`
- **`SyncUIDebugStamp`**：刷新 `bin\Release\UI` 的 HTML `?v=` 戳

---

## 4. ⚠️ HTML 版本戳（WebView2 缓存陷阱）

WebView2 用 `SetVirtualHostNameToFolderMapping` 提供 UI 且未设禁用缓存头，
不刷新 `?v=` 戳会导致用户长期加载**缓存的旧 JS**，表现为"修了却不生效"。

- 早期 `SyncUIDebugStamp` target 仅 `Condition="'$(Configuration)' == 'Debug'"` 触发，
  **Release 发布不会刷戳**——这是历史坑，现已放宽到 `Debug Or Release`，且显式传
  `--dest $(OutputPath)UI` 避免刷错到 `bin\Debug`。
- 上线后若手工拷贝/未走 MSBuild，务必用 `release-check.ps1` 复核各宿主 HTML 戳唯一且一致。

---

## 5. 重新生成 Word / PPT 技能分桶

修改任何 Skill 的 `hosts:` 后必须重跑（Excel 走主程序集 `SkillBridge`，无需生成）：

```powershell
python Hosts\Shared\Config\gen_host_skills.py
```

生成 `Hosts/Shared/Config/word_skills.json` 与 `ppt_skills.json`，再由 `SyncConfigFiles`
同步到各宿主 `bin\Release\Config`。

> 强约束：含 `executionMode=spec` 且依赖 `ExcelBridge` 的技能（如表格美化）**严禁**进
> Word/PPT 桶，否则非 Excel 宿主会渲染 Excel VBA 卡死。脚本已内置该判定。

---

## 6. MCP 工具一致性 & 质量闸门

- 工具元数据改动后：`Tools\check-tools-consistency.ps1`（VB 注册名 = JSON = 前端显示名）。
- 前端静态检查：`node Tools\lint.js`（CI 硬闸门）。
- 自动化测试：`Tools\run-tests.ps1`（JS 套件 + VB Core 测试），必须全绿。

---

## 7. ClickOnce 发布

在 Visual Studio 中用"发布"向导生成可分发安装包，发布目录为
`PublishUrl = bin\Release\`。若需自动更新，确认 `UpdateEnabled`、`UpdateMode` 配置符合预期。

---

## 8. 运行时依赖与验证

- **桥接**：VSTO 内前端走 `App.excelBridge` COM 直连，**无需手动启动桥**；
  但 `SyncBridgeFiles` 会把 `py-excel-bridge.py` + UI 部署到本机 `LocalAppData`
  （供独立网页版 / 部分能力使用）。确认安装包是否包含、用户机是否需要。
- **Python 引擎**：若启用 Python 数据分析模板，目标机需预装 Python 3.10 + `pywin32`。
- **验证必须完全退出 Office**：`.NET` 程序集进程内一次性加载，仅关工作簿不够。
  部署/验证前结束 `excel.exe / winword.exe / powerpnt.exe`，重启后新 DLL 才生效。

---

## 9. 法务与数据说明

安装包需能让用户读到：

- `Legal/EULA.md` 用户协议
- `Legal/PRIVACY.md` 隐私说明（写明云端模型会把输入和选区发到用户填写的 API 地址，本机 Ollama 默认不出本机，企业可改私有化端点）
- `Legal/THIRD_PARTY_NOTICES.md` 第三方组件

应用内：设置页有同样的风险说明；第一次发送对话前要确认一次；帮助「关于」可导出不含密钥的诊断文本。

### 9.1 席位订阅（BYOK）

订阅卖的是 **软件席位**，不计量模型 Token。授权服务见 `LicenseServer/`（默认 `http://localhost:9530`）：

```powershell
cd LicenseServer
copy .env.example .env
npm install
npm start
```

插件设置 →「账户 / 订阅」登录/试用；Stripe 与国内年付结账在系统浏览器完成。生产环境用 `LicenseSetApiBaseUrl` / 配置项 `LICENSE_API_BASE_URL` 指向公网 HTTPS。

---

## 10. 上线前人工待办清单

- [ ] 已替换为正式代码签名证书（非临时）
- [ ] 已填写 Excel / Word 的 `InstallUrl`（公网 HTTPS）
- [ ] 已把 `Legal\` 中的「【运营方】」换成营业执照主体
- [ ] 已确认 `bin\Release` 与 `publish` 中没有 `llm_config.json` 明文密钥和 `.pfx`
- [ ] 已递增 `ApplicationVersion` / `ApplicationRevision`
- [ ] 已为 Excel / Word / PPT 三宿主分别 Release 构建
- [ ] 已重跑 `gen_host_skills.py` 生成技能分桶
- [ ] 已用 `release-check.ps1` 复核（证书 / 版本戳 / 一致性 / lint / 测试）
- [ ] 已用 VS 发布向导生成 ClickOnce 安装包
- [ ] 已确认目标机 Python + pywin32（若启用模板）
- [ ] 已完全退出 Office 进程后再验证新 DLL 生效
- [ ] **服务端**：已用 `node Tools/package-server.js` 生成上传包，并确认产物里没有 `.env` / `config.json` / `data/` / `logs/`
- [ ] **服务端**：服务器上的 `.env` 已按 `.env.example` 填写 `ADMIN_ACTIVATE_KEY`（弱值等于任何人可改支付配置）

---

## 11. 服务端 / 客户端文件边界（上传清单）

仓库里同时混着「要上传到公网服务器的代码」和「只能留在用户机器上的客户端代码/密钥」，
手工挑文件极易漏。本节给出边界，并用脚本 + 守栏测试强制。

### 11.1 三类文件

| 类别 | 位置 | 去向 |
|---|---|---|
| **服务端（上传）** | `LicenseServer/server.js`、`lib/`、`public/`、`tools/`、`package.json`、`package-lock.json`、`.env.example` | 上传到公网服务器 |
| **客户端（随插件分发，绝不上服务器）** | 全部 VSTO 代码：`*.vb`、`Bridges/`、`Hosts/`、`UI/`、`Config/`、`Core/`、`Utils/`、`Resources/`、`Legal/`、`bin/` | 打进 ClickOnce 安装包 |
| **本地运营（不上公网）** | `OpsConsole/`（只绑 127.0.0.1） | 留在运营本机 |

### 11.2 一键打包（推荐，别手工挑）

```bash
node Tools/package-server.js               # 产出 dist/server/，可直接整个上传
node Tools/package-server.js --dry-run     # 只看清单，不写盘
node Tools/package-server.js --include-data # 迁移台账/用户库时（含用户数据，谨慎）
```

**永不进包**：`.env`、`config.json`（运营台写入的支付配置）、`data/`（用户库 + 密码哈希 + 台账）、
`logs/`、`node_modules/`、`test/`、任何 `.bak*` 备份。
出库前自动扫描真实密钥（`sk_live_` / `sk_test_` 长串 / `whsec_` 长串 / PEM 私钥），命中即报错退出。
`.env.example` 里的占位值（`sk_test_xxx`）不算命中，可安全随包上传作模板。

上传到服务器后：

```bash
npm install --omit=dev     # node_modules 不进包，服务器现装
cp .env.example .env       # 再按真实值填写
npm start
```

### 11.3 密钥边界（最容易出错的地方）

- **必须只在服务端**：`STRIPE_SECRET_KEY`、`STRIPE_WEBHOOK_SECRET`、`WECHAT_API_V3_KEY`、
  `WECHAT_PRIVATE_KEY`、`ALIPAY_PRIVATE_KEY`。这些一旦进了客户端 DLL，任何人都能反编译出来伪造支付。
  守栏测试 `Tests/server-package-boundary.test.js` 会扫描 `UI/js`、`Bridges`、`Hosts/Shared`，
  发现支付私钥即失败。
- **按设计在客户端**：离线激活码的 HMAC 密钥（`Bridges/LicenseKey.vb`、
  `Hosts/Shared/HostLicense.vb`、`.env` 的 `LICENSE_KEY_SECRET` 可覆盖）。
  离线验签必须用同一把对称密钥，无法避免——代价是理论上可被逆向出来自造激活码，
  这是「离线激活」方案的固有取舍，不是配置错误。
- 改 `LICENSE_KEY_SECRET` 会**作废此前发出的所有激活码**，需三处同步（两个 .vb + `.env`）并重新编译。

### 11.4 部署后自检

```bash
curl https://你的域名/health
# 期望：stripeReady / cnReady 按你在运营台填的配置返回 true/false
```

运营台填支付配置 → 写 `LicenseServer/config.json` → 热加载（需 `ADMIN_ACTIVATE_KEY` 与
`LICENSE_ADMIN_URL` 端口匹配），插件端购买按钮随之亮起。
