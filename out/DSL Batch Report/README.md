
**Script Name**: DSL Batch Report  
**Author**: Brandon Johnson  
**Date**: 2025-08-21  
**Version**: 1.1.0  

**Description**:   
Automates preperation of DSL CWR batch validation report.

**Features**:   
- Inserts columns and headers for Pref Code, AKAs, ISRCs, Setup Note, Assigned To, IPI, and PRO
- Cleans/trims song titles
- Performs lookups against "ISRCs for upload", "AKA Table", and "IP Table"
- Generates Setup Notes based on AD/AE columns
- Applies conditional formatting to highlight non-empty cells
- Formats headers with wrap, bold, underline, and fill color
- Formats additional sheets as tables and adds filter buttons
- Adds "Setup" sheet

**Notes**:   
- Column layout in main sheet matches the expected pre-insert positions
 **- The following sheets exist in the workbook**:   
     "ISRCs for upload" (lookup column )
     "AKA Table"        (lookup column A)
     "IP Table"         (lookup key column A, IPI in D, PRO in E)
- Column indexes in this script are 0-based for Office Scripts API
