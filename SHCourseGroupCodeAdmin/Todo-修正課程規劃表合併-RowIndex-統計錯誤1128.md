# Todo-修正課程規劃表合併 RowIndex 統計錯誤1128.md

## 🎯 目標
修正 `frmGPlanConfig108_MergeSubject.cs` 中新增科目筆數計算錯誤問題：  
**同樣 RowIndex 下的多筆 Subject，合併後應只算 1 筆。**  
修正完成後，請將變更記錄到：  
**課程規劃表合併調整1128.md**

## 🧩 問題描述（給 Cursor 參考）
目前程式使用：

```csharp
AddSubejctCount = AddSubjectList.Count;
```

此計算方式使用 **科目筆數**，導致：

- **同一 RowIndex 底下多個科目 → 筆數被多算**
- **沒有差異時仍顯示新增 1 筆**
- **明明應該顯示 0 筆卻出現 1 筆**

期望行為：  
**統計「不同 RowIndex 的個數」作為新增筆數。**

## ✅ Todo

### Todo 1 — 建立 RowIndex 記錄集合
在 `List<XElement> AddSubjectList = new List<XElement>();` 下方加入：

```csharp
// 用來記錄新增的 RowIndex，避免相同 RowIndex 被計算多次
HashSet<int> newRowIndexSet = new HashSet<int>();
```

### Todo 2 — 在產生新科目時記錄 RowIndex
在 `AddSubjectList.Add(NewElm);` 下方新增：

```csharp
// 記錄此 RowIndex（同一 RowIndex 只會存在一次）
newRowIndexSet.Add(LastRowIdx);
```

### Todo 3 — 修改新增筆數邏輯（核心修正）
將：

```csharp
AddSubejctCount = AddSubjectList.Count;
```

改為：

```csharp
AddSubejctCount = newRowIndexSet.Count;
```

### Todo 4 — 檢查顯示訊息是否使用 AddSubejctCount
如有使用 AddSubjectList.Count 的地方，一併改為 AddSubejctCount。

### Todo 5 — 測試
- 測試同一 RowIndex 多個科目 → 新增筆數應為 1  
- 無差異 → 新增筆數應為 0  
- 多個 RowIndex → 筆數應為 RowIndex 組數  

### Todo 6 — 修改完成後
將修正內容寫入：**課程規劃表合併調整1128.md**
