**Script Name**: Tab Export to PDF  
**Author**: Brandon Johnson  
**Date**: 2026-01-11  
**Version**: 1.0.0  
**Description**:   
Exports each worksheet in the active workbook to a PDF file after applying
basic formatting to columns H and I. Output files are saved alongside the
workbook.
**Features**:   
- Deletes column I when present.
**- Sets column widths and auto-fits columns H**: I and row heights.  
- Sanitizes worksheet names for safe PDF filenames.
- Exports every sheet as a separate PDF to the workbook folder.
**Notes**:   
- Workbook must be saved before running to determine the output folder.
- Column formatting is applied to each worksheet before export.
