Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmRoadData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim FootConversion As Decimal
	Dim InchConversion As Decimal
	Private Sub comRoadPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comRoadPrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmRoadData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmRoadData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		Dim baseunit As String
		Dim baselength As Short
		Dim i As Short
		
		If IsHelpOn = True Then
			txtRoadValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			WhichScreen = Road
			
			Call drawthevalues()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtRoadValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmRoadData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim x As Short
		
		Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		DoNotChange = True
		
		If UnitType = Metric Then
			For x = 0 To 22 Step 2
				labRoadUnits(x).Text = "meters"
				labRoadUnits(x + 1).Text = "centimeters"
			Next x
			FootConversion = 0.3048
			InchConversion = 2.54
		Else
			For x = 0 To 22 Step 2
				labRoadUnits(x).Text = "feet"
				labRoadUnits(x + 1).Text = "inches"
			Next x
			FootConversion = 1
			InchConversion = 1
		End If
		
		WhichSegment = 0
		optSegment(WhichSegment).Checked = True
		txtSegmentLabel.Text = SegNamie(WhichSegment)
		
		If PageChange(WhichScreen) = True Then
			Call drawthevalues()
		End If
		
		DoNotChange = False
		Call screenstuff()
	End Sub
	'UPGRADE_WARNING: Event frmRoadData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmRoadData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmRoadData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		On Error Resume Next
        'Me.Close()
		Call InputMenuAccess(3)
	End Sub
	Private Sub imgBackToMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgBackToMenu.Click
		Me.Close()
		Call InputMenuAccess(3)
	End Sub
	Private Sub labBackToMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labBackToMenu.Click
		Me.Close()
		Call InputMenuAccess(3)
	End Sub
	Private Sub labRoadHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labRoadHelp.Click
		Dim StartHelp As Short
		Dim SendHelp As Short
		StartHelp = 203
		Select Case WhichCell
			Case 0, 2, 4, 6, 8, 10
				SendHelp = 0
			Case 1, 3, 5, 7, 9, 11
				SendHelp = 1
			Case 12, 14, 16, 18, 20, 22
				SendHelp = 2
			Case 13, 15, 17, 19, 21, 23
				SendHelp = 3
		End Select
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(StartHelp, SendHelp)
		frmSurfaceHelp.Show()
	End Sub
	'UPGRADE_WARNING: Event optSegment.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optSegment_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSegment.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optSegment.GetIndex(eventSender)
			Dim x As Short
			WhichSegment = Index
			Call drawthevalues()
			For x = 0 To 5
				labSegment(x).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
			Next x
			labSegment(WhichSegment).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
			txtSegmentLabel.Text = SegNamie(WhichSegment)
			txtRoadValues(WhichCell).Focus()
		End If
	End Sub
	'UPGRADE_WARNING: Event txtRoadValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtRoadValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRoadValues.TextChanged
		Dim Index As Short = txtRoadValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Private Sub txtRoadValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRoadValues.Enter
		Dim Index As Short = txtRoadValues.GetIndex(eventSender)
		Dim x As Short
		WhichCell = Index
		System.Windows.Forms.SendKeys.Send("{HOME}+{END}")
		WhichCell = Index
		Call drawthevalues()
	End Sub
	Public Sub screenstuff()
		
		Dim p As Decimal
		Dim q As Decimal
		Dim r As Decimal
		Dim s As Decimal
		Dim t As Decimal
		Dim u As Decimal
		Dim v As Decimal
		
		Dim x As Short
		
		Dim y As Decimal
		Dim z As Decimal
		
		Dim h As Short
		Dim w As Short
		
		h = 6420
		w = 9150
		
		p = (300 / w) * TempWide
		q = (2160 / w) * TempWide
		
		r = (1200 / h) * TempHigh
		s = (3000 / h) * TempHigh
		t = (1020 / h) * TempHigh
		
		v = (420 / h) * TempHigh
		
		y = (480 / h) * TempHigh
		z = (660 / w) * TempWide
		
		For x = 0 To 2
			labRoadHeading(x).Top = VB6.TwipsToPixelsY((TempHigh * (180 / h)) + (x * y))
			labRoadHeading(x).Left = VB6.TwipsToPixelsX((TempWide * (180 / w)) + (x * z))
			labRoadHeading(x).Width = VB6.TwipsToPixelsX(TempWide * (1245 / w))
		Next x
		
		For x = 0 To 1
			labRoadLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (2220 / h)) + (x * r))
			labRoadLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (120 / w))
			labRoadLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (2655 / w))
			labRoadLabels(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (180 / h)) + (x * s))
			labRoadLabels(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (2880 / w))
			labRoadLabels(x + 4).Top = VB6.TwipsToPixelsY(TempHigh * (360 / h))
			If UnitType = Metric Then
				If x = 0 Then
					labRoadLabels(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (4920 / w))
					labRoadLabels(x + 4).Width = VB6.TwipsToPixelsX(TempWide * (1500 / w))
				Else
					labRoadLabels(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (6840 / w))
					labRoadLabels(x + 4).Width = VB6.TwipsToPixelsX(TempWide * (1920 / w))
				End If
			Else
				If x = 0 Then
					labRoadLabels(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (5040 / w))
					labRoadLabels(x + 4).Width = VB6.TwipsToPixelsX(TempWide * (1260 / w))
				Else
					labRoadLabels(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (6840 / w))
					labRoadLabels(x + 4).Width = VB6.TwipsToPixelsX(TempWide * (1920 / w))
				End If
			End If
		Next x
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2520 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (600 / w)) + (x * p))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2760 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (600 / w)) + (x * p))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (3720 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (480 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 22 Step 2
			If x < 12 Then
				u = 0
			Else
				u = (240 / h) * TempHigh
			End If
			LabRoadTitles(x / 2).Top = VB6.TwipsToPixelsY((TempHigh * (720 / h)) + (((x / 2) * v) + u))
			LabRoadTitles(x / 2).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			If UnitType = Metric Then
				LabRoadTitles(x / 2).Left = VB6.TwipsToPixelsX(TempWide * (3120 / w))
				txtRoadValues(x).Left = VB6.TwipsToPixelsX(TempWide * (4860 / w))
				txtRoadValues(x + 1).Left = VB6.TwipsToPixelsX(TempWide * (6900 / w))
				labRoadUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (5760 / w))
				labRoadUnits(x + 1).Left = VB6.TwipsToPixelsX(TempWide * (7800 / w))
			Else
				LabRoadTitles(x / 2).Left = VB6.TwipsToPixelsX(TempWide * (3240 / w))
				txtRoadValues(x).Left = VB6.TwipsToPixelsX(TempWide * (5040 / w))
				txtRoadValues(x + 1).Left = VB6.TwipsToPixelsX(TempWide * (7080 / w))
				labRoadUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (5940 / w))
				labRoadUnits(x + 1).Left = VB6.TwipsToPixelsX(TempWide * (7980 / w))
			End If
			txtRoadValues(x).Top = VB6.TwipsToPixelsY(TempHigh * (690 / h) + (((x / 2) * v) + u))
			txtRoadValues(x).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			txtRoadValues(x + 1).Top = VB6.TwipsToPixelsY(TempHigh * (690 / h) + (((x / 2) * v) + u))
			txtRoadValues(x + 1).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labRoadUnits(x).Top = VB6.TwipsToPixelsY((TempHigh * (720 / h)) + (((x / 2) * v) + u))
			labRoadUnits(x + 1).Top = VB6.TwipsToPixelsY((TempHigh * (720 / h)) + (((x / 2) * v) + u))
		Next x
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (4260 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (240 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (240 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (4560 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (2880 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (2940 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (2940 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (480 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (3120 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (2940 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (2940 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (3480 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (6540 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (6540 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (300 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (3180 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (6540 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (6540 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (3300 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (5940 / h))
		
		LineVertical(4).X1 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(4).X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		LineVertical(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comRoadPrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comRoadPrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labRoadHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labRoadHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	
	Public Sub drawthevalues()
		
		Dim i As Short
		Dim x As Short
		
		DoNotChange = True
		
		For i = 0 To 23
			If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
				txtRoadValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtRoadValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			Select Case i
				Case 0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22
					txtRoadValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * FootConversion)), "#,###,###,##0")
				Case Else
					txtRoadValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * InchConversion)), "##,###,##0.0")
			End Select
		Next i
		
		DoNotChange = False
		
	End Sub
	
	Private Sub txtRoadValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtRoadValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtRoadValues.GetIndex(eventSender)
		
		If KeyCode = 45 Then
			If InsertFlag = True Then
				InsertFlag = False
				labInsert.Text = "Typeover"
			Else
				InsertFlag = True
				labInsert.Text = "Insert"
			End If
		End If
		
		If InsertFlag = False Then
			Select Case KeyCode
				Case 48 To 57, 190
					If KeyCode = 190 Then
						If InStr(txtRoadValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtRoadValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRoadValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtRoadValues.GetIndex(eventSender)
		If KeyAscii > Asc("9") And KeyAscii <> Asc(",") And KeyAscii <> Asc(".") And KeyAscii <> Asc("$") Then
			Beep()
			KeyAscii = 0
		Else
			CellValues(WhichScreen, WhichCell, WhichSegment).Changed = True
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtRoadValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRoadValues.Leave
		Dim Index As Short = txtRoadValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
		WhichCell = Index
		Call Inputer(WhichCell)
	End Sub
	Private Sub Inputer(ByRef Sample As Short)
		Dim x As Short
		Dim i As Short
		Dim life As Decimal
		Dim tempvalue As String
		Dim Digit As New VB6.FixedLengthString(1)
		On Error Resume Next
		If DoNotChange = True Then Exit Sub
		PageChange(WhichScreen) = True
		tempvalue = ""
		For i = 1 To Len(txtRoadValues(Sample).Text)
			Digit.Value = Mid(txtRoadValues(Sample).Text, i, 1)
			Select Case Digit.Value
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
					tempvalue = tempvalue & Digit.Value
			End Select
		Next i
		Select Case Sample
			Case 0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					If FootConversion <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / FootConversion
				End If
			Case Else
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					If InchConversion <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / InchConversion
				End If
		End Select
		Call drawthevalues()
	End Sub
	'UPGRADE_WARNING: Event txtSegmentLabel.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSegmentLabel_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSegmentLabel.TextChanged
		If DoNotChange = True Then Exit Sub
		SegNamie(WhichSegment) = txtSegmentLabel.Text
	End Sub
End Class