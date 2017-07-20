Attribute VB_Name = "data_generator"
Option Explicit

Dim row As Integer
Dim total_rows As Integer

Sub data_generator()
Attribute data_generator.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim column As String
    
    ' If we want to fill 202 rows
    total_rows = 202
    
    ' If we want to fill in the column A
    column = "A"
    ' with string "test"
    Dim content As String
    content = "test"

    Call fill(column, 3, total_rows, content)
    
End Sub

Sub fill(ByVal column As String, ByVal begin_row As Integer, ByVal total_rows As Integer, ByVal content As String)
    For row = begin_row To total_rows
        Range(CStr(column) + CStr(row)).Value = content
    Next row
End Sub