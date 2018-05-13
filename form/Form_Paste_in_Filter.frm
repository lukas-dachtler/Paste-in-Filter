VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Paste_in_Filter 
   Caption         =   "Paste in Filter"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "Form_Paste_in_Filter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_Paste_in_Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Lbl_Rel_Click()
    Rb_Rel.Value = True
    Img_Rel.Visible = True
    Img_Abs.Visible = False
    Img_Val.Visible = False
End Sub

Private Sub Lbl_Abs_Click()
    Rb_Abs.Value = True
    Img_Rel.Visible = False
    Img_Abs.Visible = True
    Img_Val.Visible = False
End Sub

Private Sub Lbl_Val_Click()
    Rb_Val.Value = True
    Img_Rel.Visible = False
    Img_Abs.Visible = False
    Img_Val.Visible = True
End Sub

Private Sub Tb_Row_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Accepts only numeric input
    Select Case KeyAscii
        Case 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub Tb_Col_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Accept only alphabetical input
    Select Case KeyAscii
        Case 65 To 90
        Case 97 To 122
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub UserForm_Initialize()
    'Clear combobox
    Cb_Worksheet.Clear
    'Set default radiobutton
    Rb_Rel.Value = False
    Rb_Abs.Value = False
    Rb_Val.Value = True
    Img_Rel.Visible = False
    Img_Abs.Visible = False
    Img_Val.Visible = True
    Dim i As Integer
    'Load current worksheets in combobox
    For i = 1 To ActiveWorkbook.Worksheets.Count
        Cb_Worksheet.AddItem (ActiveWorkbook.Worksheets(i).Name)
    Next
End Sub

Private Sub Lbl_Paste_Click()
    'Call main function with copy option
    If Cb_Worksheet.Text = "" Or Tb_Row.Text = "" Or Tb_Col.Text = "" Then
        Dim tmp As Byte
        tmp = MsgBox("Please fill out every field!", vbCritical, "Error")
    Else
        If Form_Paste_in_Filter.Rb_Rel = True Then
            Call Module_Paste_in_Filter.Paste_in_Filter("rel")
        ElseIf Form_Paste_in_Filter.Rb_Abs = True Then
            Call Module_Paste_in_Filter.Paste_in_Filter("abs")
        Else
            Call Module_Paste_in_Filter.Paste_in_Filter("val")
        End If
    End If
End Sub

Private Sub Lbl_Cancel_Click()
    Unload Me
End Sub
