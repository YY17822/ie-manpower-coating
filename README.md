# IE 人力需求程式（Coating）

IE 部門每月月底計算未來六個月人力需求的小工具（Coating（鍍膜）版，從 Dipping_Milling 版複製來改；Dipping_Milling 版本身是從 3F 取像版複製來改）。

> 目前步驟內容還和 Dipping_Milling 版一樣，之後一步一步改成鍍膜的內容；下面各步驟的說明在改到那一步時再更新。

原本的說明：一步一步做，目前有第 1 到第 7 步（原本 3F 取像版的需求工時匯總、RMA 需求工時、RMA人力需求三步已刪除；第 7 步總表還有幾列沒填）。

名詞：**行** = 直的（A、B、C…），**列** = 橫的（1、2、3…）。

## 怎麼開啟

- **最簡單**：點兩下 `index.html`，用 Chrome 或 Edge 開啟。
- **要安裝成 App、離線使用**：點兩下 `啟動小程式.command`，瀏覽器會開啟 `http://localhost:8770`。用完把跳出來的終端機視窗關掉。

第一次使用需要連上網路，小程式會載入 ExcelJS（讀寫 Excel 的元件）。
要在不能連外部網站的電腦上用，請看 `vendor/README.txt`。

## 第 1 步：篩選 MPS 報表（Coating，版本 2 起）

可以上傳多個 MPS 報表（例如塑膠一個檔、玻璃一個檔），每個檔案各自選篩選方式和要保留的行，把列分成 Plastic、Glass 兩段，匯出成一個新的 Excel。

1. 上傳 MPS 的 Excel（只支援 `.xlsx`）。可以一次選多個，也可以分幾次加入；「往上移」調整順序、「移除」拿掉。
2. 每個檔案選一種篩選方式：
   - **依製程（Plastic）**：在料號那一行（預設 B 行，可以改）自動找 `TTL (Plastic Lens-Coating)`，上面那一塊算 Plastic。
   - **依製程（Glass）**：找 `TTL (GMO-Coating)`、`TTL (Filter&UVIR-Coating)`，上面那兩塊算 Glass（先 GMO、再 Filter&UVIR）。只找到其中一個時會提醒，只用找到的那一塊。
   - **依 Status**：輸入 A 行等於多少代表 Plastic、等於多少代表 Glass（多個值用逗號分開；這個檔案只有一種時，另一個留空白）。
   - 每個製程最後一列料號的下面是 `x`，再下面是 `TTL (製程名)`；這兩列不會留下。畫面會列出每個檔案裡找到的製程和列數。
3. 每個檔案輸入要保留的行，例如 `B, C, U, Z, AE` 或 `U:Z`，也可以點預覽表格上方的行號。
   位置可以不同，但**每個檔案保留的行數要一樣、順序要對得上**（第 1 個保留的行都放料號……）。
4. 按「匯出 Excel」，得到 `檔案1的檔名_篩選.xlsx`。

匯出的檔案：

| 列 | 內容 |
| --- | --- |
| 第 1–4 列 | 表頭，照檔案 1 的（行寬也照檔案 1） |
| 第 5 列 | 標題 `Plastic（Plastic Lens-Coating）`，跨所有保留的行置中、粗體 |
| 接著 | Plastic 的列：檔案 1 的在前，依檔案順序接下去 |
| 空 10 列 | 留給第 2 步的匯總表格 |
| 接著 | 標題 `Glass（GMO-Coating、Filter&UVIR-Coating）`，下面是 Glass 的列（同樣依檔案順序） |

- 某一段沒有任何列時會提醒，那一段只有標題。
- 表頭格式、合併儲存格的處理、公式換成數值、拿掉填滿色彩、不帶凍結窗格，都和原本一樣。
- 「匯出」下面的「製程名稱不一樣時」可以改三個製程名稱（比對不分大小寫，空白和橫線有沒有都可以），標題會跟著變。
- 每個檔案的「報表版面不一樣時」可以改資料起始列和 Status 在哪一行。
- 設定照檔案順序記住：下個月第 1 個檔案會用上次第 1 個檔案的設定。移除檔案時它的設定移到最後面，不會丟掉。
- 第 2 步以後找分段是看第 1 行的 `Plastic（…）`、`Glass（…）` 標題（`js/excel-io.js` 的 `SECTION_TITLE_RE`）。

> 原本的第 6 步（需求工時匯總）和第 7 步（RMA 需求工時）已經刪除，後面的步驟往前補成第 6 到 8 步。第 7 步總表已經改成 Coating 的「光鍍人力需求」版面（還有幾列沒填）。原本的第 8 步 RMA人力需求也已經刪除，總表是最後一步。

## 第 2 步：匯總

Plastic 和 Glass 各自在最後一列下面空一列，加一列各月的預計總產量（Total）。

1. 第 1 步的 Excel：剛做完第 1 步的話會直接沿用；也可以上傳這一版第 1 步匯出的檔案。
2. 確認各月產量在哪幾行（預設是第 5 到 10 行；第一個是當月）。這組設定和第 5 步共用。
3. 按「匯出 Excel」，得到 `…_篩選_匯總.xlsx`。

結果（以當月在 E 行、Plastic 在第 6 到 8 列、Glass 在第 20 到 31 列為例）：

每個月都只加總該月的產量；當月也不加 WIP（版本 24 起）。

| 位置 | 內容 |
| --- | --- |
| D10 | `Total` |
| E10 到 J10（各月） | `=SUM(E6:E8)`、`=SUM(F6:F8)` … |
| D33、E33… | Glass 的 Total，例如 `=SUM(E20:E31)` |

- 位置：每一段最後一列之後空一列；「Total」在第一個月左邊那一行，置中。
- 外觀：整列淺綠色底（`#B5E48C`）；「Total」加上各月圍成一塊，外框粗線、裡面細格線。
- 分段是看第 1 步寫在第 1 行的標題（`Plastic（…）`、`Glass（…）`）找出來的：標題下面、第 1 行（料號）有值的列就是那一段的資料，遇到空白列就結束。
- 某一段沒有料號時不加 Total，畫面會提醒。
- 對已經有 Total 的檔案重做這一步，會覆蓋原本那一列。
- 沒有前三大、Top 3／Non-Top 3，也不標淺橘色。

## 第 3 步：匯入 P003 標準工時

把 P003 報表（公司所有產品型號的標準工時）放進第 1 步完成的 Excel，成為同一個檔案裡名為 `P003` 的工作表。

1. 第 2 步的 Excel：剛做完第 2 步的話會直接沿用，不用先匯出；也可以上傳之前第 2 步匯出的檔案。跳過第 2 步的話，會沿用第 1 步篩選好的結果。
2. 上傳 P003 報表（`.xlsx`）。檔案裡有多個工作表時可以選要用哪一個。
3. 按「匯出 Excel」，得到 `…_匯總_含P003.xlsx`，裡面有原本的 MPS 工作表和 `P003` 工作表。

P003 工作表是整張照抄：值、格式、合併儲存格、行寬、列高、凍結窗格、自動篩選都會帶過去，公式換成數值。
上傳的檔案裡如果已經有 `P003` 工作表，會換成這次上傳的。

## 第 4 步：帶入標準工時 CT.

拿 MPS 工作表的料號到 P003 工作表找同一個料號，用那一列算出兩個單件工時：

| 新增的行 | 算法（P003 的行） |
| --- | --- |
| 機台 CT. | Machine ÷ Base Qty |
| 目檢 CT. | Labor ÷ Base Qty |

P003 的 Machine、Labor 是做 Base Qty 個的工時（Base Qty 有 1 也有 100），除過之後才是單件工時。

1. 第 3 步的 Excel：剛做完第 3 步的話會直接沿用；也可以上傳之前第 3 步匯出的檔案（要有 MPS 和 P003 兩個工作表）。
2. 選 MPS 的料號在哪一行（預設第 1 行）。
3. 選 P003 的料號、Machine、Labor、Base Qty 各在哪一行（程式會先依表頭文字帶入，請確認；P003 同時有 Machine 和 Machine/pcs 時會先選 Machine）。
4. 輸入 Plastic、Glass 各自的目檢-檢驗率（輸入 90 或 90% 都代表 90%，會記住上次的值）。
5. 按「匯出 Excel」，得到 `…_含P003_CT.xlsx`。

結果：

- 三行並排，放在 MPS 工作表「原本最後一行之後空一行」的位置：機台 CT.、目檢 CT.、目檢-檢驗率；表頭在第 2 到 4 列（合併）。
- 目檢-檢驗率是數值（不是公式）：Plastic 每一個料號都填 Plastic 的檢驗率，Glass 每一個料號都填 Glass 的；顯示成 %（例如 90.0%）。第 5 步的目檢需求工時會再乘上它。
- 只填 Plastic、Glass 兩段的料號列；第 1 步的標題、中間的空白列、第 2 步的 Total 列都不填。
- 填入的是公式，例如機台 CT. `='P003'!N5/'P003'!K5`、目檢 CT. `='P003'!M5/'P003'!K5`；在 Excel 點儲存格就能看到是從 P003 哪兩格算出來的。
- 同一個料號在 P003 出現很多列時，取最上面那一列（和 VLOOKUP 一樣）；各列算出來不一樣的料號，畫面上會列出來。
- Base Qty 是空白或 0 的料號，Excel 會顯示 `#DIV/0!`，畫面上會列出是哪些料號。
- P003 找不到的料號，兩行都留空白並標上淺紅色底，畫面上會列出是哪些料號。
- 料號比對不分大小寫、會去掉頭尾空白；`0012` 和 `12` 視為不同。
- 對已經有機台 CT.、目檢 CT.、目檢-檢驗率的檔案重做這一步，會覆蓋原本那三行，不會再多。

## 第 5 步：計算需求工時

把每個料號各月的計畫產量乘上第 4 步的「機台 CT.」算出機台需求工時；乘上「目檢 CT.」再乘「目檢-檢驗率」算出目檢需求工時。

1. 第 4 步的 Excel：剛做完第 4 步的話會直接沿用；也可以上傳之前第 4 步匯出的檔案。
2. 確認各月產量在哪幾行（預設是第 5 到 10 行）。這組設定和第 2 步共用。機台 CT.、目檢 CT. 會自動找表頭。
3. 輸入各月的表頭，兩塊用同一組。第一個（當月）用下拉選單選 Jan～Dec，再按最後一個框後面的「自動帶出後面月份」，後面幾個月會依序填好（Dec 之後接 Jan）；後面的月份也可以自己改。
4. 按「匯出 Excel」，得到 `…_CT_需求工時.xlsx`。

結果（以當月在 E 行、機台 CT. 在 L 行、目檢 CT. 在 M 行、目檢-檢驗率在 N 行為例）：

| 位置 | 內容 | 當月 | 其他月份 |
| --- | --- | --- | --- |
| O 行 | 空一行 | | |
| P 到 U 行 | 機台需求工時 | `=E6*L6` | `=F6*L6` |
| V 行 | 空一行 | | |
| W 到 AB 行 | 目檢需求工時 | `=E6*M6*N6` | `=F6*M6*N6` |

- 只算 Plastic、Glass 兩段的料號列；標題、中間的空白列、第 2 步的 Total 列都不填。
- 每一塊在 Plastic、Glass 各自最後一列下面空一列，加一列 Total（和第 2 步的 Total 同一列）：
  「Total」放在這一塊左邊那一行（O、V 行），各月 `=SUM(P6:P8)`；淺綠底（`#B5E48C`），外框粗線、裡面細線。
- 每格都是公式；當月和其他月份一樣，只用該月產量乘 CT.，不加 WIP（版本 24 起）。
- 表頭：第 2 列是跨所有月份置中的「機台需求工時」／「目檢需求工時」；第 3、4 列合併，寫你輸入的月份。
- 數值只顯示整數。儲存格裡仍保留小數，加總用的是完整的數值。
- 沒有 CT. 的料號（第 4 步在 P003 找不到）一樣會填公式，結果是 0，並標上淺紅色底；畫面上會列出是哪些料號。
- 數值的單位和 CT. 相同，不換算。
- 對已經有這兩塊的檔案重做這一步，會覆蓋原本的位置。

## 第 6 步：匯入人力報表

把人力報表（未來六個月的人力預計）放進前面做好的 Excel，成為同一個檔案裡名為 `人力` 的工作表，
並確認之後總表要抓的資料在哪裡。

1. 第 5 步的 Excel：剛做完第 5 步的話會直接沿用；也可以上傳之前第 5 步匯出的檔案。
2. 上傳人力報表（`.xlsx`）。檔案裡有多個工作表時可以選要用哪一個（有叫「人力」的會先選它）。
3. 確認六個項目各在第幾列：Online、Online（生產）、Online（重工）、Off line（直接）、Off line（間接）、TTL（On + Off直接）；
   以及當月（第一個月）在哪一行。六個月是從當月那一行開始往右連續六行。
4. 按「匯出 Excel」，得到 `…_需求工時_人力.xlsx`。

- 選單裡每一列會顯示「第 N 列　這一列的文字」。程式會先依文字帶入（例如有 `Online` 和 `生產` 的那一列就是 Online（生產）），請對照預覽確認，不對就改。
- 當月那一行：選好的各列裡的數字剛好落在六個連續的行時才會自動帶入；報表連過去的月份也列出來時，程式不知道哪一行是當月，要自己選。
- 預覽表格裡螢光的格子，就是之後總表要抓的資料；下面的核對表列出六個項目 × 六個月讀到的數值，空白或不是數字的格子標成淺紅色。
- 兩個項目選了同一列、或有空白的格子，會提醒但不擋匯出；六個項目和當月都要選好才能匯出。
- 六個項目都選好時，各項目的列會記在這台電腦的瀏覽器裡；下個月列的文字對不上時，就用記住的列當預設值。
- `人力` 工作表是整張照抄（值、格式、合併儲存格、行寬、列高），公式換成數值。上傳的檔案裡如果已經有 `人力` 工作表，會換成這次上傳的。
- 確認的位置放在 `IEApp.state.withManpower`（`rows`：各項目的列號，`firstMonthCol`：當月的行號，`monthCount`：6），給之後的總表用。

## 第 7 步：總表（光鍍人力需求，版本 4 起）

在同一個檔案裡新增 `光鍍人力需求` 工作表，版面、格式、文字照範本（標題「一般鍍膜」），並填入畫面上輸入的數值和對得上的公式。其他格之後再一項一項補上。

1. 第 6 步的 Excel：剛做完第 6 步的話會直接沿用；也可以上傳之前第 6 步匯出的檔案。
2. 工作時間：輸入工作時數、休息時數（小時，六個月共用）。
3. 各月的數值：工作天數 - 常日班、工作天數 - 做四休二、機台每日可動工時（小時）、班別數、出勤率、作業效率。
   機台每日可動工時、班別數、出勤率、作業效率只要填第一個月，再按「套用到全部月份」。
4. 人機比：GMO 站、Plastic 站各輸入機台數、人員數（六個月共用）。
5. 按「匯出 Excel」，得到 `…_人力_總表.xlsx`。

GMO = 第 1 步的 Glass 那一段（GMO-Coating + Filter&UVIR-Coating）。

目前會填的列：

| 總表的列 | 填入 | 例子 |
| --- | --- | --- |
| 工作天數 常日班、做四休二（第 5、6 列） | 各月輸入的天數 | `22` |
| 工作時間（第 7 列） | 公式（工作時數 − 休息時數）× 60；Notes 寫輸入的時數 | `=(12-1)*60` |
| 機台每日可動工時（第 8 列） | 公式 輸入的小時 × 60 | `=22*60` |
| 班別數、出勤率、作業效率（第 9–11 列） | 各月輸入；出勤率、作業效率輸入 95 或 95% 都存成 0.95，顯示 95% | `0.95` |
| MPS GMO、Plastic（第 12、13 列） | 參照 MPS 工作表 Glass、Plastic 段第 2 步的 Total 列 | `='MPS'!E25`、`='MPS'!E10` |
| GMO 站、Plastic 站 機台需求工時（第 14、19 列） | 參照該段「機台需求工時」第 5 步的 Total 列 | `='MPS'!P25`、`='MPS'!P10` |
| GMO 站、Plastic 站 人機比（第 15、20 列） | 機台數 ÷ 人員數；Notes 寫「1 人顧 4 機」 | `=4/1` |
| 目檢站 需求工時 - GMO、- Plastic（第 32、33 列） | 參照該段「目檢需求工時」第 5 步的 Total 列 | `='MPS'!W25`、`='MPS'!W10` |

- 沒有 Total 列（跳過第 2 或第 5 步）時，改成直接加總該段，例如 `=SUM('MPS'!E6:E8)`。
- 開機數、加班天數、需求操機人員、共用操機人員、Loading Rate、需求／現有人力等其他列先留空白。
- 工作天數每個月不同，不會記住；其他輸入會記住。沒填的格子留空白，畫面會提醒但不擋匯出。

版面（B 到 K 行、第 2 到 49 列，照範本）：

| 位置 | 內容 |
| --- | --- |
| 第 2 列 | 大標題「一般鍍膜」，B:K 合併置中，微軟正黑體 28 粗體，外框粗線 |
| 第 4 列 | 表頭，藍底（`#2E75B6`）白字：B:C 合併畫斜線；D 是 Unit；E 到 J 是六個月（參照 MPS 工作表「機台需求工時」的月份）；K 是 Notes |
| 第 5–11 列 | 工作天數（常日班、做四休二）、工作時間、機台每日可動工時、班別數、出勤率、作業效率 |
| 第 12–13 列 | MPS：GMO、Plastic，淺藍底 |
| 第 14–18 列 | GMO 站：機台需求工時、人機比、開機數、加班天數、需求操機人員，淺綠底 |
| 第 19–23 列 | Plastic 站：同上，淺紫底 |
| 第 24–26 列 | 共用操機人員：清洗、光譜檢測、大清機 |
| 第 27–29 列 | 需求操機人員(GMO+Plastic+共用)、現有操機人員、操機人力 Loading Rate（藍字） |
| 第 30–31 列 | 機台 Loading Rate：GMO、Plastic，淺灰底，**隱藏**（範本也是隱藏的） |
| 第 32–41 列 | 目檢站：需求工時、加班天數、需求人力（GMO、Plastic、備料）；需求目檢人員、現有目檢人員、目檢人力 Loading Rate（藍字），淺褐底 |
| 第 42–49 列 | 需求人力（On Line）、（Off Line）、總需求人力（淺橘）、現有人力（On Line）、（Off Line）、總現有人力（淺橘）、人力需求差異、總人力 Loading Rate（淺橘） |

- 字型 Calibri 12（項目、單位粗體；Notes 10），外框粗線、裡面細線，不顯示格線，顯示比例 69%；各列的數字格式照範本。
- 版面的設定都在 `js/excel-io.js` 的 `DM_REPORT`、`DM_FILL`、`DM_ROWS`、`DM_GROUPS`、`DM_UNIT_MERGES`；各列的 key（例如 `gmoMachines`）之後填內容時用。
- 對已經有 `光鍍人力需求` 或舊版 `總表` 的檔案重做這一步，會把它換成新建的。

### MPS 工作表的外觀（Dipping_Milling）

第 1、2、4、5 步在 MPS 工作表上產生的每一塊表格，外觀都統一由程式設定，不照抄原報表的樣式：

- 每一塊：外框粗線、裡面細線。
  - 第 1 步：表頭（第 2 到 4 列）一塊；Cutting、Dipping 的資料各一塊；分段標題不加框。
  - 第 2 步：每一段的 Total 列一塊（淺綠底）。
  - 第 4 步：機台 CT.、目檢 CT.、目檢-檢驗率三行合成一塊：表頭一框、每一段一框；CT. 顯示到小數三位，檢驗率顯示成 %。
  - 第 5 步：機台需求工時、目檢需求工時各一塊：表頭（大表頭加月份）一框、每一段一框；只顯示整數。
- 字型：全部用 MPS 料號那一格的字型和大小，黑色字（原報表的藍字等顏色都拿掉）。
- 找不到 CT. 的料號，淺紅色底保留。
- 匯出時會關掉 Excel 的「公式不一致」綠色小三角（每一列的公式參照不同列，Excel 會誤以為打錯）。`#DIV/0!` 這類錯誤照樣會顯示。
  做法在 `js/excel-io.js` 的 `addIgnoredErrors()`：在每張工作表的 XML 加上 `<ignoredErrors>`。

### 已知的 ExcelJS 問題：整張表變成粗體＋斜體＋刪除線

P003 這類系統匯出的報表，字型設定會把「沒有斜體、沒有刪除線」明寫成 `<i val="0">`、`<strike val="0">`。
ExcelJS 4.x 只看到標籤就當成「有」，結果匯出後整張表被畫上刪除線。
`js/excel-io.js` 的 `fixFontFlags()` 會在讀檔前先把這些標籤拿掉，所有上傳的檔案都會經過這道修正。

## 檔案

```
index.html                  主畫面
js/config.js                專案代號、Supabase 網址和公開鑰匙（SUPABASE_PUBLIC_URL、SUPABASE_PUBLISHABLE_KEY）
js/cloud.js                 登入與共用設定（Supabase）
manifest.webmanifest        安裝成 App 用的資訊
sw.js                       離線快取
css/style.css               樣式
js/app.js                   外框：步驟清單、切換步驟、載入 ExcelJS
js/excel-io.js              所有 Excel 讀寫與排版（ExcelJS）
js/ui.js                    各步驟共用的畫面元件：上傳區、試算表預覽、下載，以及月份行的共用設定
js/steps/step1-mps-filter.js    第 1 步　篩選 MPS 報表
js/steps/step2-mps-totals.js    第 2 步　匯總
js/steps/step3-p003-import.js   第 3 步　匯入 P003 標準工時
js/steps/step4-ct-lookup.js     第 4 步　帶入標準工時 CT.
js/steps/step5-demand-hours.js  第 5 步　計算需求工時
js/steps/step6-manpower.js      第 6 步　匯入人力報表
js/steps/step7-report.js        第 7 步　總表
vendor/supabase.js          supabase-js 2.117.3（連 Supabase 用，必要）
vendor/                     ExcelJS 本機副本（可選）
icons/                      圖示
啟動小程式.command           在 Mac 上啟動
.gitignore                  不進 git 的檔案（.DS_Store、.env）
```

新增步驟時：在 `js/steps/` 加一個檔案、在 `js/app.js` 的 `STEPS`、`STEP_FILES`（和要顯示「上次由誰修改」的 `STEP_SETTINGS`）加上它、
在 `sw.js` 的 `SHELL` 加上檔名並把 `VERSION` 加 1。
檔名開頭的數字就是步驟編號；在中間插入新步驟時，後面的檔案要跟著改名。

每次改版：把 `js/excel-io.js` 裡的 `VERSION` 和 `index.html` 裡每個 `?v=` 的數字一起加 1。
畫面右上角會顯示「版本 N」，可以用來確認瀏覽器跑的是不是最新的程式。

每一步做完會把結果放在 `IEApp.state`（第 1 步是 `mpsFiltered`，第 2 步是 `mpsTotals`，第 3 步是 `withP003`，第 4 步是 `withCT`，第 5 步是 `withDemand`，第 6 步是 `withManpower`，第 7 步是 `withReport`），
裡面的 `build()` 會產生一份新的活頁簿，下一步拿來接著加工。

## 登入與共用設定（Supabase）

部門的人共用同一份「記住的設定」：誰最後存，大家打開看到的就是誰的。

- 打開時要先登入（email + 密碼）。帳號由管理者在 Supabase 後台建立（Authentication → Users → Add user），不開放自行註冊。
- 設定存在 Supabase 專案的 `ie_manpower.settings` 表：一筆是 `(app, key) → value`。
  - `app`：專案代號（`js/config.js` 的 `PROJECT_ID`）。8 個人力需求 PWA 共用這張表，用 `app` 分開。
  - `key`：「步驟.輸入項目」，例如 `step7.workHours` 是第 7 步的工作時數。組名一律用步驟編號（`step1`～`step7`），第 2、5 步共用的月份行是 `step2-5`（同一個軟件的兩步共用，不是各產線共用）。
  - `value`：那個輸入項目的值（JSON）。6 個月、Off line 各種人數這種一起看的值放在同一筆（陣列）。
  - 一個輸入項目一筆，所以兩個人同時改同一步的不同項目，不會互相蓋掉；每個項目都記得是誰改的。
- 誰改的、什麼時候改的由資料庫自動記錄（`updated_by`、`updated_at`），各步驟標題下面會顯示「共用設定：上次由 某某 於 10/19 16:05 修改」。
- 右上角顯示登入的人、同步狀態（已同步／同步中／離線）和「登出」。
- 讀設定一律讀這台電腦的本機備份（localStorage），所以離線也能用；改了設定先存本機，再在背景上傳（停止輸入 1 秒後）。離線時改的會等連上網路自動上傳。
- 打開時先把 Supabase 上最新的設定抓下來蓋掉本機備份，才載入各步驟（所以 `index.html` 裡沒有各步驟的 `<script>`，改由 `js/app.js` 的 `STEP_FILES` 載入）。
- 第一次登入時，這台電腦原本有、但 Supabase 上還沒有的設定會傳上去。
- 舊版的組名第一次打開新版時會自動改成步驟編號：`totals`→`step2`、`columns`→`step2-5`、`step3`→`step4`、`step4`→`step5`、`manpower`→`step6`、`report-dm`→`step7`。從 3F 取像複製時留下、已刪除步驟的舊資料（`step5`、`step6`、`report`、`rma-manpower`）直接丟掉，不會和新的第 5、6 步撞名。

存進資料庫的輸入項目（目前和 Dipping_Milling 版相同）：

| key | 對應的輸入框 |
| --- | --- |
| `step1.files` | 第 1 步：每個檔案一組（照上傳順序）：篩選方式 `mode`（`plastic`／`glass`／`status`）、`plasticValues`、`glassValues`、保留的行 `cols`、`startRow`、`selCol`、`pnCol` |
| `step1.plasticProc`、`step1.gmoProc`、`step1.filterProc` | 第 1 步：三個製程名稱 |
| `step2.startRow` | 第 2 步：資料起始列（報表版面不一樣時） |
| `step2-5.monthCols` | 第 2、5 步共用：各月在哪幾行（`step2-5.wipCol` 是舊的 WIP 行，版本 24 起不再使用） |
| `step4.keyCol`、`step4.refKeyCol`、`step4.refMachineCol`、`step4.refLaborCol`、`step4.refBaseCol`、`step4.startRow` | 第 4 步：MPS 料號、P003 料號／Machine／Labor／Base Qty 的行、資料起始列 |
| `step4.cutRate`、`step4.dipRate` | 第 4 步：Cutting、Dipping 的目檢-檢驗率 |
| `step5.startRow` | 第 5 步：資料起始列 |
| `step6.rows`、`step6.lastFirstCol` | 第 6 步：各項目在人力報表第幾列、上次的當月行 |
| `step7.workHours`、`step7.restHours` | 第 7 步：工作時數、休息時數 |
| `step7.machineHours`、`step7.shifts`、`step7.attendance`、`step7.efficiency` | 第 7 步：機台每日可動工時、班別數、出勤率、作業效率，各 6 個月 |
| `step7.gmoMachines`、`step7.gmoPeople`、`step7.plasticMachines`、`step7.plasticPeople` | 第 7 步：GMO 站、Plastic 站的機台數和人數 |
- 不記住的欄位（第 7 步的工作天數 常日班、做四休二）照舊每次輸入，不會上傳。
- 程式在 `js/cloud.js`，用 supabase-js 套件連線（`vendor/supabase.js`，版本 2.117.3，放在專案裡，離線也載得到）。雲端、本機自架、公司主機上的 Supabase 都用同一套程式，換資料庫只改 `js/config.js` 的 `SUPABASE_PUBLIC_URL`、`SUPABASE_PUBLISHABLE_KEY`。各步驟用 `IEApp.settings.read('step7')`／`IEApp.settings.write('step7', 物件)` 讀寫一整步的設定，`cloud.js` 負責拆成一個個輸入項目上傳、下載時再組回去。
- Supabase 的網址和公開鑰匙在 `js/config.js`。公開鑰匙本來就是公開用的；安全靠登入和資料表的權限規則（只有登入的人能讀寫，不能刪除）。`SUPABASE_SECRET_KEY`（舊式名稱：`SERVICE_ROLE_KEY`）不可以放進程式。
- 取名原則：`js/config.js` 的設定名稱照自架 Supabase 的 `.env` 取，兩邊找得到同一個名字。

  | `js/config.js` | 對應自架 `.env` | 說明 |
  | --- | --- | --- |
  | `SUPABASE_PUBLIC_URL` | `SUPABASE_PUBLIC_URL` | Supabase 網址 |
  | `SUPABASE_PUBLISHABLE_KEY` | `SUPABASE_PUBLISHABLE_KEY`（舊式名稱：`ANON_KEY`） | 公開鑰匙，`sb_publishable_` 開頭 |
  | `SUPABASE_SCHEMA` | `PGRST_DB_SCHEMAS` 清單裡的其中一個 | 這個程式用的 schema（`ie_manpower`） |

- 登入資訊由 supabase-js 存在這台電腦的 `ie-manpower-coating._auth`，會自己換新憑證。
- 有登入過的人打開時，如果 4 秒內連不上 Supabase，就先用這台電腦的設定，連上之後自動同步。

### 換一個 Supabase（本機自架、公司主機）

每個新的 Supabase 都要做一次：

1. 執行下面的 SQL（Studio 的 SQL Editor）。
2. 讓 API 看得到 `ie_manpower` schema：
   - 雲端：Settings → Data API → Exposed schemas 加上 `ie_manpower`。
   - Mac 上的 supabase CLI：`supabase/config.toml` 的 `[api]` 裡 `schemas` 加上 `"ie_manpower"`，再 `supabase stop` / `supabase start`。
   - Docker 自架：docker 資料夾 `.env` 的 `PGRST_DB_SCHEMAS` 加上 `,ie_manpower`，再 `docker compose up -d` 重啟。
3. 關閉自行註冊：CLI 是 `config.toml` 的 `[auth] enable_signup = false`；Docker 是 `.env` 的 `DISABLE_SIGNUP=true`。
4. 在 Studio 的 Authentication → Users 新增同事的帳號（email + 密碼，勾 Auto Confirm）。
5. 改 `js/config.js` 的 `SUPABASE_PUBLIC_URL`、`SUPABASE_PUBLISHABLE_KEY`（各環境要填什麼寫在檔案開頭）。

搬到公司主機時要注意：

- 別台電腦連的網址要寫主機的 IP 或名稱，不能寫 `127.0.0.1`／`localhost`（那是指使用者自己的電腦）。
- 用 https 網址打開的程式（例如 GitHub Pages）不能連 http 的 Supabase（瀏覽器會擋）。程式和 Supabase 要都是 http（例如程式也放在公司主機上），或都是 https。
- 離線快取（Service Worker）和「安裝成 App」只在 https 或 `localhost` 打開時才有；用 `http://主機IP` 打開時程式照常能用，但沒有離線功能。

建立資料庫的 SQL（本機 Supabase 已經建好，8 條產線共用，不用再執行；換新的 Supabase 時才用；Exposed schemas 要加上 `ie_manpower`，見上面第 2 點）：

```sql
create schema if not exists ie_manpower;
grant usage on schema ie_manpower to authenticated;
create table if not exists ie_manpower.settings (
  app text not null, key text not null, value jsonb not null,
  updated_at timestamptz not null default now(), updated_by text,
  primary key (app, key)
);
alter table ie_manpower.settings enable row level security;
grant select, insert, update on ie_manpower.settings to authenticated;
create policy "settings_read"   on ie_manpower.settings for select to authenticated using (true);
create policy "settings_insert" on ie_manpower.settings for insert to authenticated with check (true);
create policy "settings_update" on ie_manpower.settings for update to authenticated using (true) with check (true);
create or replace function ie_manpower.stamp_settings() returns trigger language plpgsql security definer set search_path = '' as $$
begin new.updated_at := now(); new.updated_by := auth.jwt() ->> 'email'; return new; end $$;
create trigger settings_stamp before insert or update on ie_manpower.settings
for each row execute function ie_manpower.stamp_settings();
```

## 專案代號與複製成新專案

這個專案的代號是 `ie-manpower-coating`（從 Dipping_Milling 版 `ie-manpower-dipping-milling` 複製來的），寫在 `js/config.js`（網頁和 `sw.js` 都讀這個檔案）。

- 離線快取的名稱：`ie-manpower-coating-v1`（代號 + 版本號，版本號在 `sw.js` 的 `VERSION`）。新版啟用時，只刪「這個代號、但不是目前版本」的舊快取，別的專案的快取不碰。
- 記住的設定的本機備份（localStorage）：`ie-manpower-coating.step1` 這種名稱，用 `IEApp.storeKey('step1')` 產生。
- 這是新專案，`LEGACY_STORE_PREFIX`、`LEGACY_CACHE_PATTERN` 都是 `null`，不會去動其他專案的設定和快取。
- 本機啟動用的網址是 `http://localhost:8770`（3F 取像是 8765、射出是 8767、Dipping_Milling 是 8768）。

複製成新專案時要改的地方：

1. `js/config.js`：`PROJECT_ID` 改成新專案的代號；`LEGACY_STORE_PREFIX`、`LEGACY_CACHE_PATTERN` 改成 `null`。
2. `manifest.webmanifest`：`name`、`short_name`（最好也換 `theme_color` 或圖示）。
3. `啟動小程式.command`：`PORT` 改成別的號碼（例如 8766），本機測試時兩個專案才不會是同一個網址。
4. 共用設定（Supabase）不用改：同一個資料庫、同一個 schema，`PROJECT_ID` 不同就不會混在一起。
5. `sw.js` 的 `VERSION`、`js/excel-io.js` 的 `VERSION`、`index.html` 的 `?v=` 可以從 1 重新開始。

## 尚未開發

- 第 7 步總表（光鍍人力需求）的開機數、加班天數、需求操機人員、共用操機人員、Loading Rate、需求／現有人力等列
