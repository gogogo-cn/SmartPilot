# SmartPilot License Server

Seat-subscription API for SmartPilot (BYOK). Charges for the **software license** only — never meters LLM tokens and never stores user API keys.

## Quick start

The test server is a separate process. The Office add-in does not start it and does not ship a button for it.

```powershell
cd LicenseServer
.\start-test.ps1
```

That binds `http://127.0.0.1:9530` only and keeps China checkout in simulate mode. Stop it with `.\stop-test.ps1`.

For a foreground session (same defaults as production `npm start`, all interfaces):

```powershell
cd LicenseServer
copy .env.example .env
npm install
npm start
```

Default port: `9530`

## Main routes

| Method | Path | Auth | Purpose |
|---|---|---|---|
| POST | `/auth/register` | — | email/phone + password → trial |
| POST | `/auth/login` | — | returns session token |
| GET | `/me/entitlement` | Bearer | plan / periodEnd / active |
| POST | `/checkout/stripe` | Bearer | Stripe Checkout (monthly/yearly) |
| POST | `/checkout/stripe/portal` | Bearer | Stripe Customer Portal |
| POST | `/checkout/cn` | Bearer | China annual order URL |
| POST | `/webhooks/stripe` | Stripe sig | renew / cancel |
| POST | `/webhooks/cn` | HMAC | mark CN order paid |
| POST | `/admin/activate` | `x-admin-key` | manual wire-transfer unlock |

## Plans

- `trial` — `TRIAL_DAYS` (default 14) from registration  
- `pro_monthly` / `pro_yearly` — paid seats  
- `team` — same as pro with `seats` > 1 (admin activate)

价格口径（2026-10-26，与 OpsConsole/config.json 一致）：专业版月付 ¥19 / 月、年付 ¥199 / 年、
团队版 ¥129 / 席位 / 年（5 人起订，运营台按 minSeats 校验）。国内支付默认年付金额
`CN_YEARLY_AMOUNT_FEN=19900`（¥199）；Stripe 金额在 Stripe 后台的价格对象里，代码只存 Price ID。

## Stripe

Set `STRIPE_SECRET_KEY`, `STRIPE_WEBHOOK_SECRET`, `STRIPE_PRICE_PRO_MONTHLY`, `STRIPE_PRICE_PRO_YEARLY`.  
Forward webhooks: `stripe listen --forward-to localhost:9530/webhooks/stripe`.

## China annual pay

Unset merchant keys → `simulate` page at `/checkout/cn/pay.html` (dev).  
Production: set `CN_PAY_PROVIDER=wechat|alipay` and merchant credentials, then replace the stub in `lib/cnPay.js` with the official SDK unified-order call. Notify URL: `/webhooks/cn`.
