Attribute VB_Name = "Module1"
Sub CopySupSerials()
'
' CopySupSerials Macro
' Copy Supplier Serial Numbers
'

'

Workbooks.Open ThisWorkbook.Path & "\RMITCN.xlsx"
    Columns("P:P").Select
    Selection.Cut
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
Workbooks.Open ThisWorkbook.Path & "\RMITImport.xlsx"
    Range("P2").Select
    ActiveSheet.Paste

End Sub

Sub LookUpCMDB()
'
' VLOOKUP All data from RMIT CMDB Spreadsheet
'
    Workbooks.Open ThisWorkbook.Path & "\RMITImport.xlsx"
    
    Range("M2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCMDB.xlsx]Page 1'!$1:$1048576,4,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[3],'[RMITCMDB.xlsx]Page 1'!R1:R1048576,4,FALSE)"
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M597"), Type:=xlFillDefault
    Range("M2:M1000").Select
    Selection.NumberFormat = "General"
    
    Range("L2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCMDB.xlsx]Page 1'!$1:$1048576,3,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[4],'[RMITCMDB.xlsx]Page 1'!R1:R1048576,3,FALSE)"
    Range("L2").Select
    Selection.AutoFill Destination:=Range("L2:L1000"), Type:=xlFillDefault
    Range("L2:L1000").Select
    Selection.NumberFormat = "General"
    
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "PCs & Monitors"
    ActiveCell.Copy
    Range("K3:K1000").Select
    ActiveSheet.Paste
    
    Range("Q2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCMDB.xlsx]Page 1'!$1:$1048576,2,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[-1],'[RMITCMDB.xlsx]Page 1'!R1:R1048576,2,FALSE)"
    Range("Q2").Select
    Selection.AutoFill Destination:=Range("Q2:Q1000"), Type:=xlFillDefault
    Range("Q2:Q1000").Select
    Selection.NumberFormat = "General"
    
    Range("Y2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCMDB.xlsx]Page 1'!$1:$1048576,5,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[-9],'[RMITCMDB.xlsx]Page 1'!R1:R1048576,5,FALSE)"
    Range("Y2").Select
    Selection.AutoFill Destination:=Range("Y2:Y1000"), Type:=xlFillDefault
    Range("Y2:Y1000").Select
    Selection.NumberFormat = "General"
    
    Range("Z2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCMDB.xlsx]Page 1'!$1:$1048576,6,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[-10],'[RMITCMDB.xlsx]Page 1'!R1:R1048576,6,FALSE)"
    Range("Z2").Select
    Selection.AutoFill Destination:=Range("Z2:Z1000"), Type:=xlFillDefault
    Range("Z2:Z1000").Select
    Selection.NumberFormat = "General"
    
    Range("W2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCMDB.xlsx]Page 1'!$1:$1048576,7,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[-7],'[RMITCMDB.xlsx]Page 1'!R1:R1048576,7,FALSE)"
    Range("W2").Select
    Selection.AutoFill Destination:=Range("W2:W1000"), Type:=xlFillDefault
    Range("W2:W1000").Select
    Selection.NumberFormat = "General"
    
    Range("X2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCMDB.xlsx]Page 1'!$1:$1048576,8,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[-8],'[RMITCMDB.xlsx]Page 1'!R1:R1048576,8,FALSE)"
    Range("X2").Select
    Selection.AutoFill Destination:=Range("X2:X1000"), Type:=xlFillDefault
    Range("X2:X1000").Select
    Selection.NumberFormat = "General"
    
    Range("AC2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCMDB.xlsx]Page 1'!$1:$1048576,9,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[-13],'[RMITCMDB.xlsx]Page 1'!R1:R1048576,9,FALSE)"
    Range("AC2").Select
    Selection.AutoFill Destination:=Range("AC2:AC1000"), Type:=xlFillDefault
    Range("AC2:AC1000").Select
    Selection.NumberFormat = "General"
    
    Range("AD2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCMDB.xlsx]Page 1'!$1:$1048576,12,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[-14],'[RMITCMDB.xlsx]Page 1'!R1:R1048576,12,FALSE)"
    Range("AD2").Select
    Selection.AutoFill Destination:=Range("AD2:AD1000"), Type:=xlFillDefault
    Range("AD2:AD1000").Select
    Selection.NumberFormat = "General"
    
    Range("AE2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCMDB.xlsx]Page 1'!$1:$1048576,13,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[-15],'[RMITCMDB.xlsx]Page 1'!R1:R1048576,13,FALSE)"
    Range("AE2").Select
    Selection.AutoFill Destination:=Range("AE2:AE1000"), Type:=xlFillDefault
    Range("AE2:AE1000").Select
    Selection.NumberFormat = "General"
    
    Range("AR2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCMDB.xlsx]Page 1'!$1:$1048576,10,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[-28],'[RMITCMDB.xlsx]Page 1'!R1:R1048576,10,FALSE)"
    Range("AR2").Select
    Selection.AutoFill Destination:=Range("AR2:AR1000"), Type:=xlFillDefault
    Range("AR2:AR1000").Select
    Selection.NumberFormat = "General"
    
    Range("AS2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCMDB.xlsx]Page 1'!$1:$1048576,11,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[-29],'[RMITCMDB.xlsx]Page 1'!R1:R1048576,11,FALSE)"
    Range("AS2").Select
    Selection.AutoFill Destination:=Range("AS2:AS1000"), Type:=xlFillDefault
    Range("AS2:AS1000").Select
    Selection.NumberFormat = "General"

End Sub

Sub LookUpSupplier()
'
' VLOOKUP All data from CompNow Supplier Spreadsheet
'
    Workbooks.Open ThisWorkbook.Path & "\RMITImport.xlsx"
    
    Range("O2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCN.xlsx]Order_Import'!$1:$1048576,16,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[1],'[RMITCN.xlsx]Order_Import'!R1:R1048576,16,FALSE)"
    Range("O2").Select
    Selection.AutoFill Destination:=Range("O2:O1000"), Type:=xlFillDefault
    Range("O2:O1000").Select
    Selection.NumberFormat = "General"
    
    Range("G2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCN.xlsx]Order_Import'!$1:$1048576,8,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[9],'[RMITCN.xlsx]Order_Import'!R1:R1048576,8,FALSE)"
    Range("G2").Select
    Selection.AutoFill Destination:=Range("G2:G1000"), Type:=xlFillDefault
    Range("G2:G1000").Select
    Selection.NumberFormat = "General"
    
    Range("H2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCN.xlsx]Order_Import'!$1:$1048576,9,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[8],'[RMITCN.xlsx]Order_Import'!R1:R1048576,9,FALSE)"
    Range("H2").Select
    Selection.AutoFill Destination:=Range("H2:H1000"), Type:=xlFillDefault
    Range("H2:H1000").Select
    Selection.NumberFormat = "General"
    
    Range("N2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCN.xlsx]Order_Import'!$1:$1048576,15,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[2],'[RMITCN.xlsx]Order_Import'!R1:R1048576,15,FALSE)"
    Range("N2").Select
    Selection.AutoFill Destination:=Range("N2:N1000"), Type:=xlFillDefault
    Range("N2:N1000").Select
    Selection.NumberFormat = "General"
    
    Range("S2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCN.xlsx]Order_Import'!$1:$1048576,18,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[-3],'[RMITCN.xlsx]Order_Import'!R1:R1048576,18,FALSE)"
    Range("S2").Select
    Selection.AutoFill Destination:=Range("S2:S1000"), Type:=xlFillDefault
    Range("S2:S1000").Select
    Selection.NumberFormat = "General"
    
    Range("AK2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCN.xlsx]Order_Import'!$1:$1048576,37,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[-21],'[RMITCN.xlsx]Order_Import'!R1:R1048576,37,FALSE)"
    Range("AK2").Select
    Selection.AutoFill Destination:=Range("AK2:AK1000"), Type:=xlFillDefault
    Range("AK2:AK1000").Select
    Selection.NumberFormat = "0.00"
    
    Range("AL2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCN.xlsx]Order_Import'!$1:$1048576,38,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[-22],'[RMITCN.xlsx]Order_Import'!R1:R1048576,38,FALSE)"
    Range("AL2").Select
    Selection.AutoFill Destination:=Range("AL2:AL1000"), Type:=xlFillDefault
    Range("AL2:AL1000").Select
    Selection.NumberFormat = "0.00"
    
    Range("AM2").Select
    ActiveCell.Formula = "=VLOOKUP(P2,'[RMITCN.xlsx]Order_Import'!$1:$1048576,39,FALSE)"
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=VLOOKUP(RC[-23],'[RMITCN.xlsx]Order_Import'!R1:R1048576,39,FALSE)"
    Range("AM2").Select
    Selection.AutoFill Destination:=Range("AM2:AM1000"), Type:=xlFillDefault
    Range("AM2:AM1000").Select
    Selection.NumberFormat = "0.00"

End Sub

Sub Finalise()
'
' Apply Final Touches
'
    Workbooks.Open ThisWorkbook.Path & "\RMITImport.xlsx"
    
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Single Drawdown"
    ActiveCell.Copy
    Range("B3:B1000").Select
    ActiveSheet.Paste
    
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    ActiveCell.Copy
    Range("A3:A1000").Select
    ActiveSheet.Paste
    
    Range("S2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="0", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

End Sub
