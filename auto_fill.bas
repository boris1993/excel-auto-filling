Attribute VB_Name = "auto_fill"
Option Explicit

Sub auto_fill()
Attribute auto_fill.VB_ProcData.VB_Invoke_Func = " \n14"
    ' define an iterator called i
    Dim i As Integer
    
    ' r is the row number you need, like A or B or C etc,
    Dim r As String
    
    Dim rule As String
    
    ' ActiveSheet.UsedRange.Rows.Count is the total used rows
    ' e.g.: You have 5 rows with data in it, then ActiveSheet.UsedRange.Rows.Count will return 5
    ' You can also specify a number here
    For i = 1 To ActiveSheet.UsedRange.Rows.Count
    
        ' First we select the cell
        Range("A" + CStr(i)).Select
        
        ' this is the actual filling rule
        ' modify to fit your need
        ' Then we write the value into it
        ActiveCell.FormulaR1C1 = "(" + CStr(i) + ",a,a)"
        
    Next i
    
End Sub
