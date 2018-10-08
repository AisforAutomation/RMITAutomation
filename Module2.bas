Attribute VB_Name = "Module2"
Sub VLOtherWBTest()
Attribute VLOtherWBTest.VB_ProcData.VB_Invoke_Func = " \n14"
'
' VLOtherWBTest Macro
'

'
    Range("B12").Select
    Windows("RMITImportMacro.xlsm").Activate
End Sub
Sub VLIWU()
Attribute VLIWU.VB_ProcData.VB_Invoke_Func = " \n14"
'
' VLIWU Macro
'

'
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(E12,'[RMITCMDB.xlsx]Page 1'!$1:$1048576,2,FALSE)"
    Range("B12").Select
End Sub
