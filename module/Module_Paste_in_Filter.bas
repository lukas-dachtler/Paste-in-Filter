Attribute VB_Name = "Module_Paste_in_Filter"
'Variables
Private cnt_rows As Long
Private cnt_cols As Long
Private target_row As Long
Private target_col As Long
Private worksheet_index As Integer
Private curr_row As Long
Private curr_col As Long
Private cell As Range
Private temp As Byte

'*******************************************
'Description: Launches macro by showing form
'*******************************************
Public Sub Start_Paste_in_Filter()
    Form_Paste_in_Filter.Show
End Sub

'**********************************
'Description: Initializes variables
'**********************************
Private Sub Initialize()
'Catch errors
On Error GoTo err
    'Get size of current selection
    cnt_rows = Selection.Rows.Count
    cnt_cols = Selection.Columns.Count
    'Get values from form
    target_row = CLng(Form_Paste_in_Filter.Tb_Row.Text)
    target_col = Range(UCase(Form_Paste_in_Filter.Tb_Col.Text) & "1").Column
    worksheet_index = CInt(Form_Paste_in_Filter.Cb_Worksheet.ListIndex) + 1
    'Initialize values
    curr_row = target_row
    curr_col = target_col
Exit Sub
'Error-Message
err: temp = MsgBox("An error has occured!", vbCritical, "Error")
End Sub

'***************************************************************************
'Description: Pastes the values from the user selection in the filtered area
'Input      : copy_option - can be  > "rel" | relative formula
'                                   > "abs" | absolute formula
'                                   > "val" | plain values
'***************************************************************************
Public Sub Paste_in_Filter(ByVal copy_option As String)
'Catch errors
On Error GoTo err
    'Call Initialization
    Call Initialize
    'Find first row which is not hidden by filter
    Do While (Worksheets(worksheet_index).Rows(curr_row).EntireRow.Hidden = True)
        curr_row = curr_row + 1
    Loop
    'Loop over every cell from the selection
    For Each cell In Selection
        'Check if values should be written in the same row
        If (curr_col > target_col + cnt_cols - 1) Then
            'Move to next row and reset column
            curr_col = target_col
            curr_row = curr_row + 1
            'Find next visible row in filtered area
            Do While (Worksheets(worksheet_index).Rows(curr_row).EntireRow.Hidden = True)
                curr_row = curr_row + 1
            Loop
        End If
        'Paste formula or value from current cell
        If copy_option = "rel" Then
            Worksheets(worksheet_index).Cells(curr_row, curr_col).Value = cell.FormulaR1C1
        ElseIf copy_option = "abs" Then
            Worksheets(worksheet_index).Cells(curr_row, curr_col).Value = cell.Formula
        Else
            Worksheets(worksheet_index).Cells(curr_row, curr_col).Value = cell.Value
        End If
        'Move to next column
        curr_col = curr_col + 1
    Next
'Success-Message
temp = MsgBox("The values have been successfully pasted into the filtered area!", vbInformation, "Success")
Exit Sub
'Error-Message
err: temp = MsgBox("An error has occured!", vbCritical, "Error")
End Sub
