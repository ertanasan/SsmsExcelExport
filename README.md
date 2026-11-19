# SSMS Excel Export

An SSMS 22 extension (VSIX) that exports every result grid of the active query window into a single `.xlsx`, one sheet per grid.

- Menu: **Tools > Export All Result Grids to Excel...** and the query editor right-click menu.
- Reads grids through SSMS's brokered services (`IQueryEditorTabDataServiceBrokered`); it does not re-run the query.
- Status: builds; not yet verified end-to-end inside SSMS.

## Build

Uses the MSBuild bundled with SSMS 22 (references SSMS assemblies directly):

```
"C:\Program Files\Microsoft SQL Server Management Studio 22\Release\MSBuild\Current\Bin\MSBuild.exe" SsmsExcelExport.csproj -restore -p:Configuration=Release
```

Output: `bin\Release\SsmsExcelExport.vsix`. Close SSMS, double-click the VSIX to install.
