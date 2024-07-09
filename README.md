# extenda-excel-import

## Description

This component is a wrapper around [ExcelJS](https://www.npmjs.com/package/exceljs) with useful tasks added

Uses [sfmm](https://npmjs.com/package/sfmm)

## Add to project

From sf/sfdx project root:

```bash
sfmm add jsmithdev extenda-excel-import -si
```

## APIs

| Syntax | Description | Usage |
| -------- | -------- | -------- |
| newWorkbook | Create a new ExcelJS workbook | this.newWorkbook() | 
| excelToObjects | Parse Excel file to record objects | await this.excelToObjects() | 
