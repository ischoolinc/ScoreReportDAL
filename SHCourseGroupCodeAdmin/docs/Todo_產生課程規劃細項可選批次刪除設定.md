## Goal

Add **"刪除" (Delete)** as an available option in the **Batch Modify** function of `frmCreateGPlanItemSetup108.cs`.

## Scope

Target file:

* `frmCreateGPlanItemSetup108.cs`

Current batch modification options are defined by `_ProcessList`.

Currently:

```csharp
List<string> _ProcessList = new List<string>() { 
     "略過"};
```

Update the batch modification options so that users can select:

* `略過`
* `刪除`

Expected result:

```csharp
List<string> _ProcessList = new List<string>()
{
    "略過",
    "刪除"
};
```

## Requirements

1. Add `"刪除"` to the existing batch modification menu.
2. Keep `"略過"` available.
3. Do not change the existing `btnBatchModify_Click` behavior.
4. Do not change the existing `menuUpdateStatus_ItemClicked` logic.
5. The selected rows should continue to update the `"處理方式"` column through the existing logic.
6. Do not change the existing `新增`, `更新`, `刪除`, or `略過` processing logic elsewhere in the form.
7. Do not change the existing difference-status calculation.
8. Do not change the subject comparison logic.
9. Do not change sorting behavior.
10. Do not change save behavior in `btnSave_Click`.
11. Do not change `dgDataCount()` statistics logic.
12. Do not change any course code, subject level, source graduation plan, credit, semester, or XML processing logic.
13. Do not refactor unrelated code.
14. Keep the modification as small and safe as possible.

## Expected UI Behavior

When the user selects one or more rows and clicks **Batch Modify**, the menu should contain:

```text
略過
刪除
```

When `"刪除"` is selected:

* Only the currently selected rows should have their `"處理方式"` changed to `"刪除"`.
* Existing `DataGridView` behavior should remain unchanged.
* Existing delete statistics should automatically reflect the change through the current `dgDataCount()` logic.
* Saving should continue to write the selected processing status back to `chkSubjectInfo.ProcessStatus` using the existing save logic.

## Important Restrictions

Do **not** modify the original business rules that determine whether a subject is initially:

* 新增
* 更新
* 刪除
* 略過

This task only adds `"刪除"` as an additional manual **Batch Modify** option.

Do not introduce new validation or automatic status conversion unless required to fix a compile error directly caused by this change.

## Verification

After modification, verify:

1. The project compiles successfully.
2. Batch Modify displays both `"略過"` and `"刪除"`.
3. Selecting multiple rows and choosing `"刪除"` changes only those selected rows.
4. `"刪除"` count is updated correctly.
5. Clicking Save preserves `"刪除"` in `ProcessStatus`.
6. Existing `"略過"` batch modification still works.
7. Existing `新增`, `更新`, and `刪除` initial calculation results are unchanged.
8. No unrelated behavior is modified.

## Completion Record

After completing and verifying the modification, create or update:

```text
產生課程規劃調整0909.md
```

Document:

* Files modified
* Exact code changes
* Reason for the change
* Verification performed
* Confirmation that unrelated processing logic was not changed
