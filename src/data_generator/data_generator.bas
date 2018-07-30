Attribute VB_Name = "data_generator"
Option Explicit

Sub data_generator()
Attribute data_generator.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim total_rows As Integer
    Dim column As String
    Dim content As String
    
    ' How many rows do you want to generate
    total_rows = 202
    
    ' Which column do you want to fill in
    column = "A"
    
    ' The content you want to fill in to the cells
    content = "aa"

    Call fill(column, 1, total_rows, content)
    
End Sub

Sub fill(ByVal column As String, ByVal begin_row As Integer, ByVal total_rows As Integer, ByVal content As String)
    Dim row As Integer
    For row = begin_row To total_rows
        Range(CStr(column) + CStr(row)).Value = content
    Next row
End Sub
