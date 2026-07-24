
**Script Name**: DSL Batch Report  
**Author**: Brandon Johnson  
**Date**: 2026-07-24  
**Version**: 1.2.1  

**Description**:   
Automates preperation of DSL CWR batch validation report.

**Features**:   
- Inserts columns and headers for Pref Code, AKAs, ISRCs, Setup Note, Assigned To, and PRO
- Adds and populates ISWC only when an "ISWC Table" sheet exists
- Trims the existing IPI column
- Cleans/trims song titles
- Performs lookups against "ISRC Table", "AKA Table", "IP Table", and optional "ISWC Table"
- Generates Setup Notes from the shifted source rule columns
- Applies conditional formatting to highlight non-empty cells
- Formats headers with wrap, bold, underline, and fill color
- Formats additional sheets as tables and adds filter buttons
- Adds "Setup" sheet

**Notes**:   
- Column layout in main sheet matches the expected pre-insert positions
 **- The following sheets exist in the workbook**:   
     "ISWC Table"      (optional; lookup key column B, ISWC in C)
     "ISRC Table"      (lookup column )
     "AKA Table"        (lookup column A)
     "IP Table"         (lookup key column A, PRO in E)
- Column indexes in this script are 0-based for Office Scripts API
