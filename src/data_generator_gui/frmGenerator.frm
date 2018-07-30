VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGenerator 
   Caption         =   "Generator"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frmGenerator.frx":0000
   StartUpPosition =   1  'Centre of the owner
End
Attribute VB_Name = "frmGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function IsLetter(strValue As String) As Boolean
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
            Case 65 To 90, 97 To 122
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
End Function

Function null_validation() As Boolean
    If Trim(tbRow.Value = "") Or Trim(tbColumn.Value = "") Then
        null_validation = False
    Else
        null_validation = True
    End If
End Function

Function input_validation() As Boolean
    If IsLetter(tbColumn.Value) = "True" And IsNumeric(tbRow.Value) Then
        input_validation = True
    Else
        input_validation = False
    End If
End Function

Private Sub btnGenerate_Click()
    If null_validation = True And input_validation = True Then
        Dim row As Integer
        Dim total_rows As Integer
        total_rows = CInt(Trim(tbRow.Value))
        Dim column As String
        column = CStr(Trim(tbColumn.Value))
        Dim content As String
        content = CStr(Trim(tbContent.Value))
        
        For row = 1 To total_rows
            Range(column + CStr(row)).Value = content
        Next row
    Else
        Dim message, title As String
        Dim msgbox_return As Integer
        message = "Invalid input"
        title = "ERROR"
        msgbox_return = MsgBox(message, 0, title)
    End If
End Sub

Private Sub btnReset_Click()
    tbColumn.Value = ""
    tbRow.Value = ""
    tbContent.Value = ""
End Sub

Private Sub tbColumn_Change()
    tbColumn.Value = UCase(tbColumn.Value)
End Sub
