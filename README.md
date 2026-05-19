# fin-reconcile 🚀

`fin-reconcile` 是一个基于 **Python** 开发的个人财务全自动对账系统。它采用**双轨混合驱动（Hybrid-Drive）**架构，旨在彻底解决手动登录网银、下载并整理各银行 Statement 的繁琐痛点。

它完美连接了云端银行底层 API 与本地高度定制化的 Excel 资产账本，同时保留了本地 CSV 手动接管机制，兼顾了极致的自动化体验与高可靠性的降级容错方案。

---

## ✨ 核心特性 (Key Features)

* **🌐 云端 API 直连 (SimpleFIN Integration)**：通过 HTTP Basic Auth 安全直连 SimpleFIN 网关（底层接入 MX/Plaid），一键拉取涵盖 Chase、BOA、Amex、Citi 等主流大行的全量实时流水，告别网银手动导出。
* **📁 混合双轨驱动 (Hybrid-Drive Engine)**：系统优先扫描本地 `CSV for validation` 文件夹。一旦发现本地 CSV 文件，将自动**强制覆盖**云端 API 数据，完美解决部分银行（如 Discover、Venmo）过渡期账单不稳定的 API 盲区。
* **🏷️ 智能指纹映射 (Dynamic Account Mapping)**：自动读取 Excel 资产表中的账户名称、卡类型与尾号（Card#/Account#），按名称长度自适应排序，精准解决形如 `Freedom` 与 `Freedom Flex` 等相似账户名的冲突漏洞。
* **⏳ 时间缓冲匹配 (Fuzzy Date Matching)**：内置 `MAX_DAYS_TOLERANCE` 容忍度引擎（默认 5 天），自动对冲因银行结算延迟（Pending 到 Posted）导致的日期漂移，精准拦截重复或漏记账目。
* **🔒 隐私与安全优先 (Privacy First)**：采用一次性 `Setup Token` 兑换机制，本地加密保存永久 `Access URL`。绝不在代码中硬编码任何敏感凭证，支持云盘（如 OneDrive）无缝跨设备同步。

---

## 🛠️ 项目结构 (Project Structure)

```text
fin-reconcile/
├── .simplefin_url          # 自动生成的本地加密 API 凭证 (已加入 .gitignore)
├── Finances.xlsx           # 你的核心 Excel 资产账本 (已加入 .gitignore)
├── main.py                 # 双轨混合对账主程序
├── query.py                # 轻量级流水快速查询器
├── run_reconciliation.bat  # Windows 一键一键对账批处理脚本
└── CSV for validation/     # 本地 CSV 降级兜底文件夹 (已加入 .gitignore)