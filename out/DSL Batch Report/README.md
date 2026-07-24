
**Script Name**: DSL Batch Report  
**Author**: Brandon Johnson  
**Date**: 2026-07-24  
**Version**: 1.2.1  

**Description**:   
Automates preperation of DSL CWR batch validation report.

**Features**:   
- Inserts columns and headers for Pref Code, ISWC, AKAs, ISRCs, Setup Note, Assigned To, and PRO
- Trims the existing IPI column
- Cleans/trims song titles
- Performs lookups against "ISWC Table", "ISRC Table", "AKA Table", and "IP Table"
- Generates Setup Notes based on AE/AF columns
- Applies conditional formatting to highlight non-empty cells
- Formats headers with wrap, bold, underline, and fill color
- Formats additional sheets as tables and adds filter buttons
- Adds "Setup" sheet

**Notes**:   
- Column layout in main sheet matches the expected pre-insert positions
 **- The following sheets exist in the workbook**:   
     "ISWC Table"      (lookup key column B, ISWC in C)
     "ISRC Table"      (lookup column )
     "AKA Table"        (lookup column A)
     "IP Table"         (lookup key column A, PRO in E)
- Column indexes in this script are 0-based for Office Scripts API
