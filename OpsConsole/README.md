# SmartPilot 运营台（OpsConsole）

本地运营工具：把「生成激活码 → 登记台账 → 标记发货 → 对账统计 → 客户档案」从命令行和 CSV 里搬到一个网页上。**零第三方依赖**（只用 node 内置模块），**只监听 127.0.0.1**。

## 启动

**一键（推荐）**：双击 `OpsConsole\start-ops-console.bat` —— 自动找 Node.js、启动服务、等待就绪、打开 http://127.0.0.1:9600 。服务已在跑就直接打开，不会重复起。关掉弹出的黑色窗口即停止服务。

```bash
node OpsConsole/server.js        # 或 npm start --prefix OpsConsole
# 打开 http://127.0.0.1:9600
```

端口被占时：`OPS_PORT=9610 node OpsConsole/server.js`。

> 直接双击 `public/index.html`（file://）也能用：页面会自动把 API 指向本机服务，服务没起时页面顶部会出现帮助条（指明一键启动器和可复制的命令）。

## 四个视图

| 视图 | 能干什么 |
|---|---|
| 仪表盘 | 本月/累计销售额、有效/待发货/退款数、30 天到期续费清单、最近订单、方案占比 |
| 激活码台账 | 搜索（key/客户/备注）、按方案/状态筛选、到期高亮、复制 key、**一键复制客户文案**、标记已发货/退款 |
| 生成激活码 | 网页表单生成（等价 `npm run gen-key`，不用记命令行），批量自动加 `+N` 后缀一客一码 |
| 客户档案 | 按 login 聚合：订单数、累计金额（剔退款）、有效 key 数、方案分布、最晚到期 |

右上角「导出 CSV」导出合并视图（UTF-8 BOM，Excel 直接打开不乱码）。

## 数据设计（重要）

```
权威账本   LicenseServer/data/issued-keys.csv   ← gen-key 命令行和运营台都只追加，不改写
运营状态   OpsConsole/data/ops-state.json       ← 待发货/已发货/已退款 + 运营备注（按 key 索引）
定价配置   OpsConsole/config.json               ← 可选，覆盖默认定价
```

- **CSV 是唯一权威**：运营台每次请求都重读 CSV——你在命令行用 `gen-key --out` 补的货，刷新页面就能看到，不需要重启。
- **状态层与账本分离**：标记已发货/退款只写 `ops-state.json`，绝不改账本，历史永远可查。
- 账本格式与 `LicenseServer/tools/gen-license-key.js` 逐字节对齐（CRLF、7 列、引号转义），有互认测试守着（`test/compat.test.js`）。

## 配置

1. **定价**（统计销售额用）：复制下面内容存为 `OpsConsole/config.json`，按你的定价改：

```json
{
  "pro_yearly": { "amount": 199, "perSeat": false, "label": "专业版 · 年付" },
  "pro_monthly": { "amount": 19, "perSeat": false, "label": "专业版 · 月付" },
  "team": { "amount": 129, "perSeat": true, "minSeats": 5, "label": "团队版（5人以上）" }
}
```

   `perSeat: true` 表示按席位计费（amount × seats）。`minSeats` 是起订席位数——团队版 5 人起订，
   发码时不足直接拒（HTTP 400，不写台账）。不建这个文件则用内置默认值（与上面一致）。

2. **客户文案里的联系方式**：启动时加环境变量 `OPS_CONTACT=微信:hellokitty360730`，一键复制的客户文案末尾会自动带上。

## 测试

```bash
npm test --prefix OpsConsole      # 或 node --test OpsConsole/test/*.test.js
```

覆盖：CSV 解析/追加与 gen-key CLI 互认、合并视图、状态层、搜索、客户聚合、统计（退款剔除）、客户文案模板、API 端到端。同时已注册进主套件 `node Tests/run-tests.js`（`Tests/ops-console.test.js` 聚合）。

## 边界

- 单人本地工具，**无登录无鉴权**——靠只绑 127.0.0.1 保证安全，不要改绑 0.0.0.0 或部署到服务器。
- 退款只做状态标记（不计入统计），账本记录保留——真要作废某个 key 需换 secret 重新编译发布（见 MEMORY.md 订阅章节）。
