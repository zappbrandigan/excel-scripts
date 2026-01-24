**Script Name**: Split Column A to Sheets  
**Author**: Brandon Johnson  
**Date**: 2026-01-11  
**Version**: 1.0.0  
**Description**:   
Splits values from column A into multiple worksheets, capped by a maximum
rows-per-sheet value. Optionally targets a specific worksheet and can skip
the header row.
**Features**:   
- Detects row count from column A or the overall used range.
- Splits data into new sheets named Split_01, Split_02, etc.
- Copies only column A values to keep output minimal.
- Auto-fits the output column for readability.
**Parameters**:   
**- maxRowsPerSheet**: number  
  **Maximum number of data rows per output worksheet (default**: 30000).  
**- sheetName?**: string  
  Optional worksheet name to target instead of the active sheet.
**- skipHeader**: boolean  
  **Whether to ignore the first row as a header (default**: true).  
**Notes**:   
- Only column A is copied to the new sheets.
- Existing worksheets are not modified; new ones are added.
