Attribute VB_Name = "Module3"
Sub Removezerostest()
Attribute Removezerostest.VB_Description = "remove zeros"
Attribute Removezerostest.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Removezerostest Macro
' remove zeros
'

'
    Range("R2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="0", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub
