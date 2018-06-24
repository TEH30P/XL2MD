# XL2MD
*Tiny script to convert Excel Workbook XLS/XLSX into Markdown document.*

Worksheets converted to separate files with name `<xls/xlsx file full name or param -iPathMD>.<worksheet name>.md`. Forbiddened in file name characters is not escaped.
Script uses microsof office excel ActiveX objects so you need to have installed office.

## Parameters
- `iPathXLS` input xls/xlsx file path.
- `iPathMD` path to resulting markdown file.
