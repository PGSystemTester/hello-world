
Private Sub ColorFooterComboBox1_Change()
Dim i As Long, strLayoutType, strSlideCount As Long, strSlideName As String



strLayoutType = ColorFooterComboBox1.ListIndex + 1
strSlideCount = ActivePresentation.Designs(strLayoutType).SlideMaster.CustomLayouts.Count

DividerLayoutComboBox1.Clear
'----- Fill in combobox  -------

For i = 1 To strSlideCount
    strSlideName = ActivePresentation.Designs(strLayoutType).SlideMaster.CustomLayouts(i).Name
    
    If InStr(LCase(strSlideName), "divider") > 0 Then

'        DividerLayoutComboBox1.AddItem Replace(strSlideName, "Divider - ", "")
        DividerLayoutComboBox1.AddItem strSlideName & getStrSeparator1 & i
    
    End If

Next i
DividerLayoutComboBox1.ListIndex = 0


'---------------------------------

End Sub



Private Sub CommandButton1_Click()

'-----input validation
If IsNumeric(MoveupTextBox) = False Or IsNumeric(ToplimitTextBox) = False Then

   MsgBox "Please enter only numeric values on Titles and Paragraphs Offset"
   Exit Sub

End If
'-------------------------

Me.Hide


End Sub


Private Sub CommandButton2_Click()
End
End Sub





Private Sub UserForm_Initialize()
Dim i As Long

'----- Fill in combobox  -------

ColorFooterComboBox1.List = Array("Light", "Dark")
ColorFooterComboBox1.ListIndex = 0

'For i = 1 To 11
'
'''DividerLayoutComboBox1.AddItem Replace(Replace(ActivePresentation.Designs(8).SlideMaster.CustomLayouts(i).Name, "Divider - ", ""), "Macro", "")
'DividerLayoutComboBox1.AddItem Replace(ActivePresentation.Designs(2).SlideMaster.CustomLayouts(i).Name, "Divider - ", "")
'
'Next i
'DividerLayoutComboBox1.ListIndex = 3


'---------------------------------

End Sub
