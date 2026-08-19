
## Goal

Enhance `frmGPlanConfig108_MergeSubject.cs` so that when subjects are merged from one curriculum plan into another, each merged `<Subject>` records which source curriculum plan it came from.

After completing the modification, document the changes and test results in:

`課程規劃表合併調整0817.md`

## Target File

`frmGPlanConfig108_MergeSubject.cs`

Current merge logic is based on:

* `SelectedGPlanNameList`
* `GPlanDict`
* `sourceGPlanInfo`
* `sourceGPlanInfo.RefGPContentXml.Elements("Subject")`
* `new XElement(elm)`
* `TargetGPlanInfo.RefGPContentXml`

The variable `name` inside:

```csharp
foreach (string name in SelectedGPlanNameList)
```

represents the name of the source curriculum plan currently being merged.

## Required Changes

### 1. Add source curriculum plan information to newly merged Subject

Locate the code where the source `<Subject>` is copied:

```csharp
XElement NewElm = new XElement(elm);
```

Immediately after creating `NewElm`, add a new attribute to record the source curriculum plan name.

Use:

```csharp
NewElm.SetAttributeValue("來源課規", name);
```

Expected result:

```xml
<Subject
    課程代碼="..."
    OfficialSubjectName="..."
    來源課規="來源課程規畫表名稱">
```

The value must come directly from the current `name` variable.

Do not hard-code curriculum plan names.

### 2. Preserve all existing Subject data

When adding `來源課規`, do not change or remove any existing `<Subject>` attributes or child nodes.

Existing data such as:

* `課程代碼`
* `OfficialSubjectName`
* `Grouping`
* `Grouping.RowIndex`
* other existing attributes/elements

must remain unchanged.

Only append the new `來源課規` attribute.

### 3. Do not overwrite existing target curriculum plan Subjects

Subjects that already exist in:

```csharp
TargetGPlanInfo.RefGPContentXml.Elements("Subject")
```

before the merge should remain unchanged.

Do not automatically add or modify `來源課規` for existing target records unless they are newly copied from a selected source curriculum plan during this merge.

### 4. Preserve existing merge rules

Do not change the existing primary merge behavior.

Keep the current logic for:

* selected source curriculum plans
* target curriculum plan selection
* course code duplicate checking
* `RowIndex` calculation
* grouping behavior
* `AddSubjectList`
* `newRowIndexSet`
* `AddSubejctCount`
* `GetTargetGPlanInfo()`
* `GetAddSubjectCount()`

The purpose of this task is only to add source curriculum plan information to newly merged `<Subject>` records.

### 5. Review duplicate CourseCode behavior

Review this existing condition:

```csharp
if (!TargetCourseCodeList.Contains(CourseCode))
```

Currently, `TargetCourseCodeList` is initially populated from the target curriculum plan.

Check whether a newly added `CourseCode` should also be added into `TargetCourseCodeList` after being accepted.

For example:

```csharp
TargetCourseCodeList.Add(CourseCode);
```

This may be necessary to prevent two selected source curriculum plans from both adding the same `CourseCode`.

Do not change this behavior blindly.

First verify whether the intended business rule is:

> The same course code should only exist once in the final merged curriculum plan.

If this is already the intended rule, update the in-memory duplicate tracking accordingly.

Document the result in the change record.

## Expected Data Flow

```text
Selected source curriculum plan
        ↓
SelectedGPlanNameList
        ↓
foreach (string name ...)
        ↓
GPlanDict[name]
        ↓
sourceGPlanInfo
        ↓
sourceGPlanInfo.RefGPContentXml
        ↓
Subject elm
        ↓
new XElement(elm)
        ↓
NewElm
        ↓
add attribute:
來源課規 = name
        ↓
AddSubjectList
        ↓
TargetGPlanInfo.RefGPContentXml
```

## Example

Source curriculum plan:

```text
113學年度普通科課程規畫
```

Original source XML:

```xml
<Subject 課程代碼="A001" OfficialSubjectName="國文">
    <Grouping RowIndex="10" />
</Subject>
```

After merging:

```xml
<Subject
    課程代碼="A001"
    OfficialSubjectName="國文"
    來源課規="113學年度普通科課程規畫">
    <Grouping RowIndex="..." />
</Subject>
```

Only the `RowIndex` may be adjusted according to the existing merge logic.

## Validation

Test at least the following cases:

* [ ] Merge one source curriculum plan into one target curriculum plan.
* [ ] Confirm every newly added `<Subject>` contains `來源課規`.
* [ ] Confirm `來源課規` equals the actual source curriculum plan name.
* [ ] Merge multiple source curriculum plans and confirm each new `<Subject>` records the correct source.
* [ ] Confirm existing target `<Subject>` records are not unexpectedly modified.
* [ ] Confirm existing `課程代碼` duplicate checking still works.
* [ ] Test two selected source curriculum plans containing the same `CourseCode`.
* [ ] Confirm `Grouping.RowIndex` behavior is unchanged.
* [ ] Confirm `AddSubejctCount` still counts unique newly added RowIndex groups correctly.
* [ ] Confirm the final merged XML structure remains valid.
* [ ] Build the project and confirm there are no compilation errors.

## Constraints

* Do not redesign the merge architecture.
* Do not modify unrelated UI behavior.
* Do not rename existing methods or variables unless necessary.
* Do not change the existing XML structure except for adding the new `來源課規` attribute.
* Keep compatibility with the current project and .NET Framework environment.
* Keep changes small and focused.

## Completion Record

After implementation and testing, create or update:

`課程規劃表合併調整0817.md`

Record:

1. Modified files.
2. Original behavior.
3. New `來源課規` behavior.
4. Exact location where the attribute is written.
5. Source of the value (`name` from `SelectedGPlanNameList`).
6. Whether duplicate `CourseCode` tracking was adjusted.
7. Test cases executed.
8. Test results.
9. Any compatibility or regression considerations.
