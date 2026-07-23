# 第六學期成績單加註 Metadata 調整（0331）

## 問題現象

執行第六學期成績單 PDF 產出並寫入 `GradeData` Metadata 時，曾出現：

`Unknown prefix. Need to call method RegisterNamespaceURI`

## 原因

舊寫法透過 **`SaveNewInfoWithXmp`**（或 XMP／prefix 相關路徑）寫入 Metadata 時，觸發 Aspose.PDF 對 XMP namespace／prefix 的處理，導致上述錯誤。

## 修正方式

改為使用 **Aspose.PDF 8.8.0**（專案：`Aspose.Pdf_201402.dll`）之 **`PdfFileInfo`**：

1. `BindPdf(pdfPath)`
2. `SetMetaInfo("GradeData", json)`（**不使用** `dc:`、`xmp:` 等含冒號的 key）
3. `SaveNewInfo(pdfPath)`（**不使用** `SaveNewInfoWithXmp`，**不使用** `Document.Metadata[...]`）

並同步以同名 `.json` 備份相同 JSON 內容（UTF-8 無 BOM）。

---

## 1. 修改目的

在 `Student6thSemesterCorseCodeRank` 產生 PDF 後，使用 **Aspose.Pdf**（專案庫 `Aspose.Pdf_201402.dll`，對應文件所述 Aspose.PDF 8.x）將第六學期成績資料寫入 PDF 自訂 Metadata 欄位 `GradeData`，並同步產出與 PDF 同主檔名之 `.json` 備份檔；不變更既有 Word／PDF 產出流程與畫面操作。

## 2. 修改檔案

- `Report/Student6thSemesterCorseCodeRank.cs`：合併結果與修課人數改為 class-level 欄位、Metadata／JSON 邏輯與批次錯誤提示。
- `SHCourseGroupCodeAdmin.csproj`：新增 Aspose.Pdf 組件參考（`Library\Aspose.Pdf_201402.dll`）。

## 3. 新增方法名稱

| 方法 | 說明 |
|------|------|
| `ValidateGradeMetaItem` | 驗證單筆 `GradeMetaItem` 欄位與格式 |
| `GetCourseStudentCount` | 依課程代碼自 `_CourseStudentCountDic` 取修課人數 |
| `GetCourseCategoryCode` | 自 `CourseCode` 解析課程類別代碼（23 碼以上） |
| `GetDomainCode` | 自 `CourseCode` 解析領域名稱代碼（23 碼以上） |
| `GetSkipLoopForSixthSemesterRow` | 與報表相同之不計學分／不需評分篩選 |
| `BuildGradeDataJson` | 依學生組出 `GradeData` JSON 字串 |
| `WriteGradeDataToPdf` | 寫入 PDF Metadata 並輸出同名 `.json` |
| `TryWriteGradeDataToPdf` | 單一學生失敗時記錄訊息，不中斷其他學生 |

內部類別：`GradeMetaItem`（七個欄位，名稱為繁體中文）。

## 4. Metadata 寫入位置

- `BgWorkerReport_RunWorkerCompleted`：在 `document.Save(..., SaveFormat.Pdf)` 成功後呼叫 `TryWriteGradeDataToPdf`。
- 另存新檔（`SaveFileDialog` 成功儲存 PDF）後同樣呼叫 `TryWriteGradeDataToPdf`。

## 5. JSON 備份檔產出規則

- 路徑：與實際輸出之 PDF 同目錄、同主檔名，副檔名改為 `.json`。
- 編碼：UTF-8（無 BOM）。
- 內容：與寫入 `GradeData` 之字串相同；最外層為 JSON 陣列，每筆固定七欄（繁中欄位名）。

## 6. 課程類別代碼／領域名稱代碼取法

- **課程類別代碼**：`CourseCode` 長度須大於 22（至少 23 碼），取 **第 17 碼**（0-based index `16`），與 `frmCheckSCAttendCourseCode` 等處對 23 碼課程代碼之用法一致；再轉大寫並驗證 `^[1-9A-F]$`。
- **領域名稱代碼**：同為 23 碼以上時，取 **Substring(19, 2)**，與 `DataAccess.CourseCodeConvertToGPlanByGroupCode`／`GPlanInfo108` 中領域對照之取碼一致。

若長度不足或無法解析，會拋出例外，不寫入非法資料。

## 7. 測試結果

- 請於開發環境執行 `msbuild` 或 Visual Studio 建置通過後，依 Todo「第十五節 測試重點」手動驗證：單一／批次學生、另存新檔、`GradeData` 與 `.json`、單一學生失敗不中斷整批、不再出現 `Unknown prefix`。

## 8. 仍需確認事項

- **課程類別代碼**取法：程式內以 23 碼以上課程代碼第 17 碼為準，註解標 `TODO: 依正式課程代碼規則取得課程類別代碼`。
- **領域名稱代碼**取法：以 `Substring(19, 2)` 為準，註解標 `TODO: 依正式課程代碼規則取得領域名稱代碼`。

## 9. 風險與注意事項

- **Aspose 授權**：`Aspose.Pdf_201402.dll` 須與主程式部署路徑一致；若執行環境未含該 DLL，Metadata 寫入會失敗。
- **課程代碼格式**：非 23 碼或第 17 碼不在 `1–9`／`A–F` 時，該生 Metadata 會失敗並列入批次錯誤訊息。
- **與報表一致**：`GradeData` 列與報表列印列使用相同篩選（含最多 60 科、排除修課中、需有分數與排名等）；若與實際紙本需求不一致，應再對照規格調整。
- **批次錯誤**：單一學生 Metadata 失敗時，其餘學生仍會產出 PDF；結束後以訊息方塊列出失敗學生與原因。
