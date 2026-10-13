# 06 · Power Apps (Power Fx)

`MonitorKlien-OnSave.fx` is the **OnSave** formula of a customised SharePoint list form (Power Apps). The form edits one parent job item and, on a second screen, an editable grid of its child step rows stored in another list.

What the formula does:

1. **Concurrency check** – compares the item's `Modified` value with the value captured when the form opened; blocks the save if someone else edited the item meanwhile.
2. Saves the parent form (`SubmitForm`) and computes progress %.
3. Splits the grid collection into **new** and **changed** rows:
   - new rows → one bulk `Patch(DataSource, ForAll(Defaults), ForAll(records))`
   - changed rows → `ForAll` + `LookUp` by ID
4. Builds Lookup / Person values with the proper `@odata.type` records (`SPListExpandedReference`, `SPListExpandedUser`).
5. Error handling with `IfError` / `FirstError`, then closes the form (`RequestHide`).

> File extension is `.fx` for readability; paste the content into the form's `OnSave` property.
