# NPI Portal — EdgeOne 部署倉庫

EdgeOne Pages 專案 `npi-check` 的部署來源：push 到 `main` 即自動部署。
本倉庫只放「要上線的檔案」，由 `NPI_vx` 資料夾的建置流程自動產生，**請勿手改**。
（手改會在下次建置時被白名單管制品清掉。）

## 結構（一功能一資料夾，首頁在根目錄）

| 路徑 | 內容 | 線上網址 |
|---|---|---|
| `index.html` | NPI 資料入口首頁（源：NPI_vx\portal.html） | https://npi-check.edgeone.dev/ |
| `NPI_vx_web/` | 查詢 + 檢查工具 | `/NPI_vx_web/npi_search.html`、`/NPI_vx_web/TTM_check.html` |
| `webnpi_web/` | 產品矩陣（源：webnpi\product_matrix.html） | `/webnpi_web/index.html` |
| `npi_dashboard*.html` / `*.xlsx` / `MP變動記錄*` | 根目錄其餘上線檔 | `/npi_dashboard.html` 等 |
| `functions/api/expect-mp.js` | Pages Function（期望MP 回存 KV，GET/POST） | `/api/expect-mp` |
| `TTM_expect_mp.json` | 期望MP 靜態基準（頁面 fallback 用） | `/TTM_expect_mp.json` |

## 更新方式（都在 `NPI_vx` 資料夾操作）

- **NPI 五件套 + 首頁**：雙擊 `一键生成2.bat`（建置 → 複製進本倉庫 → push）
- **產品矩陣**：雙擊 `push_webnpi.bat`（只推 `webnpi_web/index.html`）
- **新增功能頁**：自足 html 放好 → 加進 `sync_web.py` 的 `MAPPING` → 首頁源 `portal.html` 加卡片 → 跑 `一键生成2.bat`

## 注意

- `functions/` **必須在倉庫根目錄**（Pages Functions 路由約定，不能收進子資料夾）
- 線上另有 EdgeOne「訪問密碼」全站保護；期望MP 回存密碼 `EXPECT_MP_TOKEN` 與 KV 綁定都是 EdgeOne 專案層級設定，不在本倉庫
- 舊倉庫 `kabonka/npi-dashboard` 是 NPI_vx 建置工作區的備份，不再負責上線
