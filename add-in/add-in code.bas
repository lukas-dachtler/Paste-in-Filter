Private Sub Workbook_Open()
Set Btn = Application.CommandBars("Tools").Controls.Add(Type:=msoControlButton, Temporary:=True)
	With Btn
		.Style = msoButtonIconAndCaption
		.FaceId = 640
		.Caption = "Paste in Filter"
		.OnAction = "Start_Paste_in_Filter"
	End With
End Sub