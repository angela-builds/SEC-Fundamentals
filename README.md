# sec-fundamental-tool

從 SEC 官方資料庫（data.sec.gov）直接抓上市公司近 10 年財報數字，整理成 Excel，
不受一般免費網站只顯示 5 年的限制，數字都能追溯回原始申報文件。

---

## 環境設定（第一次使用，或換一台電腦時）

```
python -m venv venv
venv\Scripts\activate          # Mac/Linux 用 source venv/bin/activate
pip install pandas numpy openpyxl yfinance
```

每次要用之前，先啟用虛擬環境（終端機最前面出現 `(venv)` 代表啟用成功）：

```
venv\Scripts\activate
```

`01_fetch_sec_companyfacts.py` 裡的 `HEADERS` / `SEC_FILES_HEADERS` 需要填入真實聯絡信箱
（SEC 規定），換電腦或分享這份專案給別人使用時記得檢查一下。

---

## 在哪裡輸入股票代碼（ticker）

**單一股票，一支一支跑**：直接把 ticker 當參數帶在指令後面，例如：

```
python scripts\01_fetch_sec_companyfacts.py NVDA
python scripts\02_extract_annual_metrics.py NVDA
python scripts\03_add_valuation.py NVDA
python scripts\04_build_summary.py NVDA
python scripts\05_build_charts.py NVDA
```

不帶參數直接執行也可以，程式會跳出來問你要查哪支股票。

**多支股票一次跑（推薦，日常維護用這個）**：

```
python scripts\run_pipeline.py AAPL PG NVDA
```

或者把想固定追蹤的股票清單寫進 `config\tickers.csv`（第一欄是 `ticker`），不帶參數直接執行：

```
python scripts\run_pipeline.py
```

它會依序處理清單裡每一支，01→05 都跑完，最後自動跑一次 06 品質檢查。想重跑但不想
再打一次 SEC API（例如只是在測試邏輯有沒有改對），加 `--skip-fetch`：

```
python scripts\run_pipeline.py AAPL PG NVDA --skip-fetch
```

跑完的 Excel 檔案在 `outputs\excel\{TICKER}_annual_fundamentals.xlsx`。

---

## 完整流程（含拆股檢查，定期上傳資料前建議照這個順序跑一次）

1. `python scripts\run_pipeline.py`　　（或帶特定 ticker）— 跑完 01-06
2. `python scripts\07_detect_stock_splits.py` — 掃描所有股票，找出可能的拆股候選，
   結果寫在 `outputs\06_quality_check\07_split_candidates.csv`，**不會**自動修改任何資料
3. **人工核對**候選清單是不是真的拆股（查一下公司公告或 8-K），確認後把資訊填進
   `config\confirmed_splits.csv`（同一支股票如果拆過不只一次，每次都加一行，見檔案裡
   AAPL / NVDA 的範例）
4. `python scripts\08_apply_split_adjustments.py` — 套用確認過的拆股調整，重建
   summary / charts_data，原始未調整數字會保留在 `_as_reported` 欄位，不會被覆蓋掉

07、08 刻意沒放進 `run_pipeline.py` 自動跑，因為「這是不是真的拆股」需要人工判斷，
不建議讓程式自己默默套用。

---

## 適用範圍（重要，先看這段再決定要不要查某支股票）

這個工具只適用於：**在美國依 US GAAP 準則對 SEC 申報財報的美國公司**（一般是用
Form 10-K 申報），且累積至少 3 年以上申報歷史。「有沒有在美股掛牌交易」不是判斷
標準，以下這幾種情況即使股票代碼看起來是美股，也不適用：

| 情況 | 例子 | 會發生什麼 |
|---|---|---|
| 外國私人發行人（用 20-F、IFRS 準則申報，不是 10-K/US GAAP） | UL、PHG、BABA、NVO | 直接報錯「No annual facts found」，這不是 bug，是這類公司的財報標籤本來就不在程式找的 `us-gaap` 命名空間底下 |
| 上市或分拆時間太短，累積不到 3 年資料 | SOLV、剛上市不久的公司 | 直接報錯「Only found N fiscal years」 |
| 累積 3-9 年，但不到 10 年 | 視個股情況 | 可以跑，但會印出警告，`summary` 分頁的 `years_covered` 欄位會標明實際涵蓋幾年，平均值欄位名稱也會反映實際年數，不會誤標成「10 年」 |
| 傳統註冊型 ETF（依 1940 年投資公司法申報 N-CSR，不是 10-K） | 大多數股票型/債券型 ETF | 通常會直接找不到資料；即使找得到，「營收」「EPS」這些概念對 ETF 本身也沒有意義，不建議用這個工具查 ETF |
| 商品信託型 ETF（例外，有申報 10-K） | ARKB 這類現貨比特幣 ETF | 技術上可能抓得到部分 `us-gaap` 資料，但財務指標對這類被動信託一樣沒有實質意義，不建議當作个股分析使用 |
| 股利分配權利依股票類別不同、須用「雙類別法」分別計算 EPS 的公司（不是單純表決權分層） | HSY（普通股/B 類股股利權利不同） | `EarningsPerShareDiluted` 這個標籤的資料在改用雙類別法申報後就斷層（HSY 是 2011 年起），因為分類後的數字帶有維度標記，companyfacts API 這層簡化資料抓不到，需要另外解析完整 XBRL 文件才能取得，目前工具不支援。**注意**：這跟「多股票類別」本身無關——Alphabet（GOOG/GOOGL，A/B/C 股表決權不同但股利權利相同）、多數只是表決權分層的公司，都只用單一 EPS 數字申報，不受影響，已實測確認可正常使用 |

簡單說：**先確認是不是一般美國公司、用 US GAAP 申報，再決定要不要查**，外國發行人
跟 ETF 目前都不在範圍內。

**Colab 使用**：打開 `colab_runner.ipynb`（用 Google Drive 或直接上傳到 Colab 開啟），
裡面已經照順序排好：裝套件 → 掛載 Drive → git clone/pull 程式碼 → 設定 SEC 聯絡信箱
→ 跑批次流程 → 拆股偵測/確認/套用 → 把結果同步回 Drive。第一次用記得把 notebook 裡
的 `REPO_URL` 換成你自己的 repo 網址。

**信箱設定**：`SEC_CONTACT_EMAIL` 這個環境變數取代了原本寫死在程式碼裡的信箱。本機
用終端機設定（每次開新的終端機視窗都要重設，或寫進系統環境變數一勞永逸）：
```
$env:SEC_CONTACT_EMAIL = "your_email@example.com"
```
Colab 裡則是 `colab_runner.ipynb` 第 4 格直接設定（不會被 commit 進 git，因為那格內容
是你自己在 Colab 介面上跑的，不是存在檔案裡）。

---

## 免責聲明

本工具數字來源為 SEC 官方申報資料，經自動化流程整理，僅供學習與討論參考，不構成
投資建議。使用者應自行查核關鍵數字並獨立判斷，作者不對資料誤差或依此做出的投資
決策負責。

---

## 檔案結構速查

```
config/
  tickers.csv            批次處理用的股票清單
  confirmed_splits.csv   人工確認過的拆股資訊（07/08 會用到）
  peer_companies.csv     同業比較用的股票清單（如果有在用這個功能）
data/raw/                每支股票的原始 SEC companyfacts JSON（01 產生）
outputs/excel/           每支股票的最終 Excel 檔案（02-05, 08 都會寫入這裡）
outputs/06_quality_check/ 品質檢查報告、拆股候選清單
outputs/logs/            批次執行（run_pipeline.py）失敗時的完整錯誤紀錄
outputs/charts/          每支股票的圖表
scripts/                 01-08 各步驟腳本 + run_pipeline.py 批次執行器
```
