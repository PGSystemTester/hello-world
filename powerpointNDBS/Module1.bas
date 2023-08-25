Private Declare PtrSafe Function WaitMessage Lib "user32" () As Long

Public Const getStrSeparator1 = "->"

Sub importPPTFunction()
    Dim filename1 As String, singlefilename As String
    Dim dividerColor As Long, layoutColor As Long, _
    srtMoveUp As Long, topLimit As Long, dividerSlideNumber As Long
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    topLimit = -10
    
    selectLayoutUserForm.MoveupTextBox = srtMoveUp
    selectLayoutUserForm.ToplimitTextBox = topLimit
    
    selectLayoutUserForm.Show vbModeless 'show UF

    Do While selectLayoutUserForm.Visible = True
      DoEvents
    Loop


    dividerColor = selectLayoutUserForm.DividerLayoutComboBox1.ListIndex + 1
    layoutColor = selectLayoutUserForm.ColorFooterComboBox1.ListIndex + 1
    dividerSlideNumber = CLng(Split(selectLayoutUserForm.DividerLayoutComboBox1.Text, getStrSeparator1)(1))


    '--- title and paragraph offset ---
    srtMoveUp = CLng(selectLayoutUserForm.MoveupTextBox)
    topLimit = CLng(selectLayoutUserForm.ToplimitTextBox)
    
    '-----------------------

    Unload selectLayoutUserForm

    filename1 = ShowFileDialog

    If filename1 = "" Then Exit Sub

'----- get single file name -------
    Dim rawFileName As String, p As Long, fileEnding As String
    rawFileName = fso.GetFilename(filename1)
    p = Len(rawFileName)
    Do While Mid(rawFileName, p, 1) <> "."
        p = p - 1
    Loop
    fileEnding = Mid(rawFileName, p + 1, 9)
    
    
    

singlefilename = Replace(Replace(rawFileName, ".pptx", ""), ".ppt", "")
singlefilename = singlefilename & "-NTTDATA.pptx"
singlefilename = ShowSaveAsDialog(singlefilename)
Stop

If singlefilename = "" Then Exit Sub


Dim pre As Presentation

Set pre = ActivePresentation


Call deleteAllSlides(pre)

Call saveAsPPT(singlefilename, pre)

Call copySlides(filename1)

Call changeLayout(layoutColor, dividerSlideNumber)

Call editShapesInteli

Call moveShapeUp(srtMoveUp, topLimit) '


Call deleteAdditionalPattern
Call deleteAdditionalPattern2

Call checkFooters

ActivePresentation.Save

MsgBox "Done"

End Sub


Sub saveAsPPT(filename1 As String, pre As Presentation)
    With pre
        .BuiltInDocumentProperties.Item("title").Value = .Name
        .SaveAs filename1
    End With
End Sub

Private Sub deleteAllSlides(pre As Presentation)
Dim x As Long

    For x = pre.Slides.Count To 1 Step -1
        pre.Slides(x).Delete
    Next x

End Sub

Private Sub copySlides(strFilename As String)
    Dim objPresentation As Presentation, thisPresentation As Presentation
    Dim strProgress As Double, initSlide As Long, _
    i As Integer, iCounter As Long
    Dim objSectionProperties As SectionProperties, newSectionProperties As SectionProperties
    
    Set thisPresentation = ActivePresentation
    'open the target presentation
    
    Set objPresentation = Presentations.Open(strFilename)
    
    If objPresentation.Slides.Count < 1 Then Exit Sub 'exit if not enought slides
    
    Set objSectionProperties = objPresentation.SectionProperties
    Set newSectionProperties = thisPresentation.SectionProperties
    
    
    If LCase(ActivePresentation.Slides.Item(1).CustomLayout.Name) = "title" Then 'set init Slide
      initSlide = 2
    Else
      initSlide = 1
    End If
    
    initSlide = 1


    On Error Resume Next
    For i = initSlide To objPresentation.Slides.Count
    
        '==== Update progress bar ===
        strProgress = i * 100 / objPresentation.Slides.Count
        ProgressBarUserForm.ProgressLabel.Caption = "Importing content, " & Round(strProgress, 1) & "%, please wait...."
        ProgressBarUserForm.ProgressBar.Width = strProgress * 2
        ProgressBarUserForm.Show
        DoEvents
        '==============
       
        objPresentation.Slides.Item(i).Copy
        
        Call Wait(1.2)
        
        thisPresentation.Slides.Paste
    
        Presentations.Item(1).Slides.Item(Presentations.Item(1).Slides.Count).Design = _
            objPresentation.Slides.Item(i).Design
    
    Next i

    '..... Adding sections as well ......
    For iCounter = 1 To objSectionProperties.Count
        newSectionProperties.AddBeforeSlide objSectionProperties.FirstSlide(iCounter), objSectionProperties.Name(iCounter)
    Next iCounter
    '........................
    
    On Error GoTo 0
    
    objPresentation.Close
    
    thisPresentation.Save
    
    Unload ProgressBarUserForm

End Sub



Private Sub changeLayout(layoutType As Long, dividerSlideNumber As Long)
Dim i As Integer, k As Integer, macroPattern As Long
Dim sourceSlideName As String, targetSlideName As String
Dim strProgress
Dim sld As Slide, shp As Shape
Dim shapeArray
Dim specTxtArray

If layoutType = 1 Then
  macroPattern = 3 'light
  
Else
  macroPattern = 4 'dark
  
End If

On Error Resume Next
For i = 1 To ActivePresentation.Slides.Count

    Set sld = ActivePresentation.Slides.Item(i)
    
    Call changeOnlyShapesColor(sld)
    
    If i > 1 Then
        Call changeOnlyTextColor(sld, layoutType)
    End If

    '==== Update progress bar ===
    strProgress = i * 100 / ActivePresentation.Slides.Count
    ProgressBarUserForm.ProgressLabel.Caption = "Checking layout, " & Round(strProgress, 1) & "%, please wait...."
    ProgressBarUserForm.ProgressBar.Width = strProgress * 2
    ProgressBarUserForm.Show
    DoEvents
    '==============

    sourceSlideName = LCase(ActivePresentation.Slides.Item(i).CustomLayout.Name)
    
    shapeArray = getShapeFeatures(sld)
    
    If InStr(sourceSlideName, "title") > 0 Or InStr(sourceSlideName, "cover") > 0 Or InStr(sourceSlideName, "titel") > 0 Then
          
          If i > 1 Then
            GoTo anotherTitle
          End If
                              
          sld.CustomLayout = ActivePresentation.Designs(layoutType).SlideMaster.CustomLayouts(1)
    
    ElseIf InStr(sourceSlideName, "agenda") > 0 Then
    
       sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(1) 'Old macro layout
                 
    ElseIf InStr(sourceSlideName, "content") > 0 Then
    
       sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(2) 'Old macro layout
       
    ElseIf InStr(sourceSlideName, "two columns") > 0 Then
       
       sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(3) 'Old macro layout
               
    ElseIf InStr(sourceSlideName, "sheer") > 0 Then
       
        sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(4) 'Old macro layout
       
    ElseIf InStr(sourceSlideName, "divider") > 0 Then  'separating slides
    
       Call deleteOnlyPicturesInSlide(sld) 'Delete Pictures on this slide

       sld.CustomLayout = ActivePresentation.Designs(layoutType).SlideMaster.CustomLayouts(dividerSlideNumber)
         
    Else
    
anotherTitle:
       sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(2) 'content Slide
       
    End If
    

changeshape:

    If i > 1 And InStr(sourceSlideName, "divider") = 0 Then
        Call getShapePositionBack(sld, sourceSlideName, shapeArray) 'change shape to its original position
    End If


    '-----  delete empty boxes ---------
    If i > 1 And sourceSlideName <> "sub title" And sourceSlideName <> "agenda" And sourceSlideName <> "1_sub title" And sourceSlideName <> "title" Then
      Call DeleteShapeWithSpecTxt(sld, "")
    End If
    '-----------------------------------
    
    '----- delete shapes with specific text -----
    specTxtArray = Array("© 2010 itelligence", "© 2011 itelligence", "© 2012 itelligence", "© 2013 itelligence", "© 2014 itelligence", "© 2015 itelligence", "© 2016 itelligence", "© 2017 itelligence", _
    "© 2018 itelligence", "© 2019 itelligence", "© 2020 itelligence", "© 2021 itelligence", "© 2022 itelligence", "© 2023 itelligence")
    For k = 0 To UBound(specTxtArray)
       Call DeleteShapeWithSpecTxt2(sld, CStr(specTxtArray(k)))
    Next k
    Call DeleteShapeWithSpecTxt2(sld, "We Transform. Trust into Value")
    '-------------------------------------------------
    
    
    '..... delete image on separate and cover slide if exist ......
    If i > 1 And sourceSlideName = "sub title" Or sourceSlideName = "1_sub title" Or sourceSlideName = "title" Then
    
      Call deleteOnlyPicturesInSlide(sld)
'      For Each shp In sld.Shapes
'         If shp.HasTextFrame = False Then
'           shp.Delete
'         End If
'      Next
    End If
    '.....................................
    
    '..... send image back last sheet .....
    If sourceSlideName = "final" Or sourceSlideName = "1_final" Then
       For Each shp In sld.Shapes
         If shp.HasTextFrame = False Then
            shp.ZOrder msoSendToBack
         End If
       Next
    End If
    '...............................
    
  
    Call changeFontBulletColor(sld, layoutType)
   
    
    targetSlideName = LCase(ActivePresentation.Slides.Item(i).CustomLayout.Name)
    Call change2FinalLayout(sld, targetSlideName, layoutType)
    
    If i > 1 And InStr(sourceSlideName, "divider") = 0 Then
        Call getShapePositionBack(sld, sourceSlideName, shapeArray) 'change shape to its original position
    End If
    
Next i

Unload ProgressBarUserForm


ActivePresentation.Save

End Sub

Sub change2FinalLayout(sld As Slide, targetSlideName As String, layoutType As Long)
Dim pivotShapeName As String, pivotShapeNameSub As String
Dim shp As Shape

On Error Resume Next
If layoutType = 1 Then
 
  macroPattern = 3
Else
  
  macroPattern = 4
End If

'----- mapp text Cover -----
If Mid(targetSlideName, 1, 5) = "cover" Then
   pivotShapeName = getLowestCustomName(sld, "Custom Shape Name")
   pivotShapeNameSub = "Custom Shape Name " & pivotShapeName + 1
   pivotShapeName = "Custom Shape Name " & pivotShapeName
   
   
   For Each shp In sld.Shapes
      If shp.HasTextFrame And InStr(shp.Name, "Custom Shape Name") <> 0 And InStr(shp.Name, pivotShapeName) = 0 And InStr(shp.Name, pivotShapeNameSub) = 0 Then
          sld.Shapes(pivotShapeNameSub).TextFrame.TextRange.Text = sld.Shapes(pivotShapeNameSub).TextFrame.TextRange.Text & vbNewLine & vbNewLine & shp.TextFrame.TextRange.Text
          shp.TextFrame.TextRange.Text = ""
      End If
   Next
   
   Call DeleteShapeWithSpecTxt(sld, "")
End If
'--------------------
      
'----- mapp text Divider -----
If Mid(targetSlideName, 1, 7) = "divider" Then
   pivotShapeName = getLowestCustomName(sld, "Custom Shape Name")
   pivotShapeName = "Custom Shape Name " & pivotShapeName
   
   For Each shp In sld.Shapes
      If shp.HasTextFrame And InStr(shp.Name, "Custom Shape Name") <> 0 And InStr(shp.Name, pivotShapeName) = 0 Then
          sld.Shapes(pivotShapeName).TextFrame.TextRange.Text = sld.Shapes(pivotShapeName).TextFrame.TextRange.Text & vbNewLine & vbNewLine & shp.TextFrame.TextRange.Text
          shp.TextFrame.TextRange.Text = ""
      End If
   Next
   
  Call DeleteShapeWithSpecTxt(sld, "")
End If
'--------------------

Select Case targetSlideName
    
    Case "agenda macro"
    
       sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(2)
       Call changeBullets12Numbers(sld)
    
    Case "content macro"
    
       sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(2)
       Call DeleteShapeWithSpecTxt(sld, "")
    
    Case "two columns macro"
       
       sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(3)
       Call DeleteShapeWithSpecTxt(sld, "")
        
    Case "sheer macro"
       
        sld.CustomLayout = ActivePresentation.Designs(macroPattern).SlideMaster.CustomLayouts(1)
        Call DeleteShapeWithSpecTxt(sld, "")
    
    End Select

End Sub


Private Function ShowFileDialog() As String
    With Application.FileDialog(Type:=msoFileDialogFilePicker)
        .Title = "Select file to convert"
        .Filters.Clear
        .Filters.Add "All PowerPoint Types", "*.pptx;*.ppt;*.pptm;*.ppsx;*.pps;*.ppsm;*.potx;*.pot;*.potm;*.odp"
        .AllowMultiSelect = False
        If .Show Then ShowFileDialog = .SelectedItems.Item(1)
    End With
End Function
Function ShowSaveAsDialog(Optional strFilename As String)

On Error Resume Next

ShowSaveAsDialog = ""

    Dim dlgSaveAs As FileDialog

    Set dlgSaveAs = Application.FileDialog(Type:=msoFileDialogSaveAs)

    With dlgSaveAs
        ''.AllowMultiSelect = False
        .InitialFileName = strFilename
        .Show
        ''ShowSaveAsDialog = dlgOpen.SelectedItems.Item(1)
        ShowSaveAsDialog = .SelectedItems.Item(1)
    End With
    
 
End Function


Sub editShapesInteli()
Dim sld As Slide
Dim i As Integer


'inches x 72 to get points


For i = 1 To ActivePresentation.Slides.Count

    Set sld = ActivePresentation.Slides.Item(i)
    
'    Call deleteShapesByPosition(sld, 0, 0) '
    Call deleteShapesByPosition(sld, 0, 520) '
    Call deleteShapesByPosition(sld, 29.43, 520.86) '
    Call deleteShapesByPosition(sld, 471.71, 520.86) '
    Call deleteShapesByPosition(sld, 856.28, 523.14) '

    
    'Call moveShapeUp(sld, 30) 'move shapes up
       
Next i




End Sub

Sub deleteShapesByPosition(sld As Slide, strLeft As Long, srtTop As Long)
Dim shp As Shape
Dim strNumber As Double

strNumber = 0.5
'inches x 72 to get points


    For Each shp In sld.Shapes
     
       If shp.Left >= strLeft * 0.95 - strNumber And shp.Left <= strLeft * 1.05 + strNumber And shp.Top >= srtTop * 0.95 - strNumber And shp.Top <= srtTop * 1.05 + strNumber Then
         
         shp.Delete
       
       End If
    Next



End Sub


Public Sub Wait(Seconds As Double)
    Dim endtime As Date
    endtime = VBA.Now + Seconds / 3600 / 24
    Do
        WaitMessage
        DoEvents
    Loop While VBA.Now < endtime
End Sub

Sub DeleteShapeWithSpecTxt(oSld As Slide, sSearch As String)
  Dim lShp As Long
  
  
  On Error GoTo errorhandler
  'If sSearch = "" Then sSearch = ActivePresentation.Slides(335).Shapes(4).TextFrame.TextRange.Text

  
  For lShp = oSld.Shapes.Count To 1 Step -1
      With oSld.Shapes(lShp)
        If .HasTextFrame And InStr(oSld.Shapes(lShp).Name, "Custom Shape") = 0 Then
          If StrComp(sSearch, .TextFrame.TextRange.Text) = 0 Then .Delete
        End If
      End With
  Next
  
Exit Sub
errorhandler:
  Debug.Print "Error in DeleteShapeWithSpecTxt : " & Err & ": " & Err.Description
  
End Sub
Sub DeleteShapeWithSpecTxt2(oSld As Slide, sSearch As String)
  Dim lShp As Long
  
  
  On Error GoTo errorhandler
  'If sSearch = "" Then sSearch = ActivePresentation.Slides(335).Shapes(4).TextFrame.TextRange.Text

  
  For lShp = oSld.Shapes.Count To 1 Step -1
      With oSld.Shapes(lShp)
        If .HasTextFrame And InStr(oSld.Shapes(lShp).Name, "Custom Shape") <> 0 Then
          ''If LCase(sSearch) = LCase(Trim(.TextFrame.TextRange.Text)) Then .Delete
          If InStr(.TextFrame.TextRange.Text, sSearch) <> 0 Then .Delete
        End If
      End With
  Next
  
Exit Sub
errorhandler:
  Debug.Print "Error in DeleteShapeWithSpecTxt : " & Err & ": " & Err.Description
  On Error GoTo 0
End Sub

Sub moveShapeUp(srtMoveUp As Long, topLimit As Long)
  Dim oSld As Slide
  Dim lShp As Long
  Dim newTop As Double
  Dim i As Long
  Dim sourceSlideName As String
  
  
  
For i = 1 To ActivePresentation.Slides.Count


  '==== Update progress bar ===
    strProgress = i * 100 / ActivePresentation.Slides.Count
    ProgressBarUserForm.ProgressLabel.Caption = "Moving up shapes, " & Round(strProgress, 1) & "%, please wait...."
    ProgressBarUserForm.ProgressBar.Width = strProgress * 2
    ProgressBarUserForm.Show
    DoEvents
    '==============

  Set oSld = ActivePresentation.Slides.Item(i)
  
  sourceSlideName = LCase(ActivePresentation.Slides.Item(i).CustomLayout.Name)
  
  If (sourceSlideName = "cover slide 1 macro" Or sourceSlideName = "cover slide 14 macro") And i = 1 Then 'do not affect cover
     GoTo skipShape
  End If
  
  
  
  For lShp = oSld.Shapes.Count To 1 Step -1
  
      If lShp < 1 Then Exit For
      
      With oSld.Shapes(lShp)
      
           If .HasTextFrame = True And (.Height >= 530 Or .Top <= topLimit) Then GoTo skipShape  '(.Top <= 30 And .Width >= 900) Or
           
           If .Top < 20 And .Width > 850 Then GoTo skipShape  'do not change  titles
           
           newTop = .Top - srtMoveUp
           
           
          .Top = newTop
          
            
          '----- set limit ----
          If .Top < topLimit And .HasTextFrame = True Then
             .Top = topLimit
          End If
          '-------------------------------
       
      End With
skipShape:
  Next

Next i

Unload ProgressBarUserForm

End Sub

Sub bringTextToFront()
  Dim oShp As Shape, oSld As Slide, i As Long
  Dim sourceSlideName As String

    For i = 1 To ActivePresentation.Slides.Count
    
    
      '==== Update progress bar ===
        strProgress = i * 100 / ActivePresentation.Slides.Count
        ProgressBarUserForm.ProgressLabel.Caption = "Bring text to front, " & Round(strProgress, 1) & "%, please wait...."
        ProgressBarUserForm.ProgressBar.Width = strProgress * 2
        ProgressBarUserForm.Show
        DoEvents
        '==============
    
      Set oSld = ActivePresentation.Slides.Item(i)
      
      sourceSlideName = LCase(ActivePresentation.Slides.Item(i).CustomLayout.Name)
      
    
          '.... bring text to front ...
        For Each oShp In oSld.Shapes
             If oShp.HasTextFrame = True Then
               If Len(oShp.TextFrame.TextRange.Text) > 1 Then
                 oShp.ZOrder msoBringToFront
               End If
             End If
        Next
        '..........................
      
    
    Next i
    
    Unload ProgressBarUserForm

End Sub

Private Function getShapeFeatures(sld As Slide)
    Dim shp As Shape
    Dim shapeArray(), counter As Long
    Dim x As Long
    
    'inches x 72 to get points
    counter = 0
    ReDim shapeArray(counter)
    shapeArray(counter) = ""
    
    ''Set sld = ActivePresentation.Slides.Item(1)
    For Each shp In sld.Shapes
      
        shp.Name = "Custom Shape Name " & counter
        ReDim Preserve shapeArray(counter)
        shapeArray(counter) = shp.Name & ";" & shp.Top & ";" & shp.Left & ";" & shp.Height & ";" & shp.Width
        'shapeArray(counter) = shp & ";" & shp.Top & ";" & shp.Left & ";" & shp.Height & ";" & shp.Width
        counter = counter + 1
      If shp.Type = msoGroup Then
         
         For x = 1 To shp.GroupItems.Count
         'MsgBox shp.GroupItems.Count
            
            shp.GroupItems(x).Name = "Custom Shape Name " & counter
            ReDim Preserve shapeArray(counter)
            shapeArray(counter) = shp.GroupItems(x).Name & ";" & shp.GroupItems(x).Top & ";" & shp.GroupItems(x).Left & ";" & shp.GroupItems(x).Height & ";" & shp.GroupItems(x).Width
            counter = counter + 1
         Next x
      End If
    Next
    
    getShapeFeatures = shapeArray
End Function

Sub changeFontBulletColor(sld As Slide, layoutType As Long)
        Dim i As Long, j As Long, k As Long
        Dim x As Long
        Dim oTbl
        Dim currentSlideName As String
        Dim rustRedRGB As Long
        
        rustRedRGB = RGB(178, 32, 0)
        
        
        currentSlideName = sld.CustomLayout.Name
        
         'Set sld = ActivePresentation.Slides.Item(3)
         
        
        For lShp = sld.Shapes.Count To 1 Step -1
              With sld.Shapes(lShp)
                If .HasTextFrame Then
                   
                   '....... change blue box that was set by mistake .......
        '           If .Fill.ForeColor.RGB = RGB(0, 128, 177) Then
        '              .Fill.ForeColor.RGB = RGB(255, 255, 255)
        '           End If
                   '........................................
                   
                   '............  Change text color and size ........
                   If Len(.TextFrame.TextRange.Text) > 1 Then
        '                .TextFrame.TextRange.Font.Name = "Arial"
                        For i = 1 To Len(.TextFrame.TextRange.Text)
                            
                            If layoutType = 1 Then 'light
                                .TextFrame.TextRange.Characters(i).Font.Color = getColorConversion(.TextFrame.TextRange.Characters(i).Font.Color.RGB)  '
                            Else
                                 .TextFrame.TextRange.Characters(i).Font.Color = getColorConversion(.TextFrame.TextRange.Characters(i).Font.Color.RGB, True) 'Dark Mode
                            End If
                           
                           
                           '/////   size  /////////
                           If Mid(LCase(currentSlideName), 1, 7) <> "divider" Then
                             If .TextFrame.TextRange.Characters(i).Font.Size > 24 Then
                                '.TextFrame.TextRange.Characters(i).Font.Size = 24
                             End If
                           End If
                           '///////////////////////
                           
                        Next
                   End If
                   '............................................
                   
                   '......... Change bullet color ..................
                   For i = 1 To .TextFrame.TextRange.Paragraphs.Count
                     If .TextFrame.TextRange.Paragraphs(i).ParagraphFormat.Bullet.Visible <> 0 Then
                        .TextFrame.TextRange.Paragraphs(i).ParagraphFormat.Bullet.Font.Color = rustRedRGB
                     End If
                   Next i
                   '.......................................................
                
                ElseIf .HasTable Then
                  Set oTbl = sld.Shapes(lShp).Table
                    For k = 1 To oTbl.Columns.Count
                      For j = 1 To oTbl.Rows.Count
                        
                        Call changeShapeColor(sld.Shapes(lShp).Table.Cell(j, k).Shape) 'Change cell color
                        
                        With oTbl.Cell(j, k).Shape.TextFrame.TextRange
                          '...... change color text in tables .......
                          '.Size = 12
        '                  .Font.Name = "Arial"
                          For i = 1 To Len(.Text)
                              .Characters(i).Font.Color = getColorConversion(.Characters(i).Font.Color.RGB) 'Change to Blue
                          Next i
                          '.Bold = True
                          '.....................................
                          
                          
                          '......... Change bullet color ..................
                          For i = 1 To .Paragraphs.Count
                              If .Paragraphs(i).ParagraphFormat.Bullet.Visible <> 0 Then
                                 .Paragraphs(i).ParagraphFormat.Bullet.Font.Color = rustRedRGB ''getColorConversion(.Paragraphs(i).ParagraphFormat.Bullet.Font.Color.RGB)  'Change to Blue
                              End If
                          Next i
                          '.......................................................
                          
                          
                        End With
                      Next j
                    Next k
                           
                   
                End If
                
                '---- shape color -----
                Dim oShpNode
                Dim oNode As SmartArtNode
                If .HasSmartArt Then
                   For Each oNode In .SmartArt.Nodes
                      For Each oShpNode In oNode.Shapes ' As ShapeRange
                         Call changeShapeColor(oShpNode)
                      Next
                   Next
                End If
                
                
                On Error Resume Next
                If .HasTable = False Then
                    If .Type <> msoGroup Then
                      Call changeShapeColor(sld.Shapes(lShp))
                    Else
            
                       'Debug.Print "GROUP"
                       For x = 1 To sld.Shapes(lShp).GroupItems.Count
                           Call changeShapeColor(sld.Shapes(lShp).GroupItems(x))
                           
                          
                           
                           '****** check for texts*****************
                           If sld.Shapes(lShp).GroupItems(x).HasTextFrame Then
                              If Len(sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Text) > 1 Then
                                 For i = 1 To Len(sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Text)
                                    sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Characters(i).Font.Color = getColorConversion(sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Characters(i).Font.Color.RGB)  'Change to Blue
                                 Next i
                              End If
                              
                              
                              '............ Change bullet color ...........................
                              For i = 1 To sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Paragraphs.Count
                                If sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Paragraphs(i).ParagraphFormat.Bullet.Visible <> 0 Then
                                   sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Paragraphs(i).ParagraphFormat.Bullet.Font.Color = rustRedRGB '
                                End If
                              Next i
                              '..........................................................
                           
                           End If
                           '***************************************
                           
                           '******* has chart ************
                           On Error Resume Next
                           If sld.Shapes(lShp).GroupItems(x).HasChart Then
                             If sld.Shapes(lShp).GroupItems(x).Chart.ChartType = 51 Then
                             
                                ''shp.Chart.SeriesCollection(1).DataLabels.Font.Color = RGB(0, 0, 0)
                                sld.Shapes(lShp).GroupItems(x).Chart.SeriesCollection(1).Interior.Color = getColorConversion(sld.Shapes(lShp).GroupItems(x).Chart.SeriesCollection(1).Interior.Color)
                                sld.Shapes(lShp).GroupItems(x).Chart.SeriesCollection(1).Border.Color = getColorConversion(sld.Shapes(lShp).GroupItems(x).Chart.SeriesCollection(1).Border.Color)
                             End If
                           End If
                           On Error GoTo 0
                           
                           '******************************
                           
                           
                       Next x
                    End If
                End If
                '--------------------------
              
              End With
        Next

End Sub
 

Private Sub changeOnlyTextColor(sld As Slide, layoutType As Long)
    Dim i As Long, j As Long, k As Long, x As Long, rustRedRGB As Long
    Dim oTbl
    Dim currentSlideName As String
    
    rustRedRGB = RGB(178, 32, 0)
    
    currentSlideName = sld.CustomLayout.Name
    
    For lShp = sld.Shapes.Count To 1 Step -1
          With sld.Shapes(lShp)
            If .HasTextFrame Then
    
               
               '............  Change text color ........
               If Len(.TextFrame.TextRange.Text) > 1 Then
    '
                    For i = 1 To Len(.TextFrame.TextRange.Text)
                        
                        If layoutType = 1 Then 'light
                            .TextFrame.TextRange.Characters(i).Font.Color = getColorConversion(.TextFrame.TextRange.Characters(i).Font.Color.RGB)  '
                        Else
                             .TextFrame.TextRange.Characters(i).Font.Color = getColorConversion(.TextFrame.TextRange.Characters(i).Font.Color.RGB, True) 'Dark Mode
                        End If
                       
                    Next
               End If
               '............................................
               
            
            ElseIf .HasTable Then
              Set oTbl = sld.Shapes(lShp).Table
                For k = 1 To oTbl.Columns.Count
                  For j = 1 To oTbl.Rows.Count
                 
                    With oTbl.Cell(j, k).Shape.TextFrame.TextRange
                      '...... change color text in tables .......
    
                      For i = 1 To Len(.Text)
                          .Characters(i).Font.Color = getColorConversion(.Characters(i).Font.Color.RGB) 'Change to Blue
                      Next i
                      '.Bold = True
                      '.....................................
                      
                    End With
                  Next j
                Next k
                       
               
            End If
            
            '---- Text color -----
            On Error Resume Next
            If .HasTable = False Then
                   'Debug.Print "GROUP"
                   For x = 1 To sld.Shapes(lShp).GroupItems.Count
                       '****** check for texts*****************
                       If sld.Shapes(lShp).GroupItems(x).HasTextFrame Then
                          If Len(sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Text) > 1 Then
                             For i = 1 To Len(sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Text)
                                sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Characters(i).Font.Color = getColorConversion(sld.Shapes(lShp).GroupItems(x).TextFrame.TextRange.Characters(i).Font.Color.RGB)  'Change to Blue
                             Next i
                          End If
                       End If
                       '***************************************
                   Next x
            End If
            '--------------------------
          
          End With
    Next
End Sub

Private Sub changeOnlyShapesColor(sld As Slide)
Dim j As Long, k As Long, x As Long, rustRedRGB As Long
Dim oTbl
Dim currentSlideName As String

rustRedRGB = RGB(178, 32, 0)

currentSlideName = sld.CustomLayout.Name

For lShp = sld.Shapes.Count To 1 Step -1
      With sld.Shapes(lShp)

        If .HasTable Then
          Set oTbl = sld.Shapes(lShp).Table
            For k = 1 To oTbl.Columns.Count
              For j = 1 To oTbl.Rows.Count
                
                Call changeShapeColor(sld.Shapes(lShp).Table.Cell(j, k).Shape) 'Change cell color
                
              Next j
            Next k
                              
        End If
        
        '---- shape color -----
        Dim oShpNode
        Dim oNode As SmartArtNode
        If .HasSmartArt Then
           For Each oNode In .SmartArt.Nodes
              For Each oShpNode In oNode.Shapes ' As ShapeRange
                 Call changeShapeColor(oShpNode)
              Next
           Next
        End If
        
        
        On Error Resume Next
        If .HasTable = False Then
            If .Type <> msoGroup Then
              Call changeShapeColor(sld.Shapes(lShp))
            Else
    
               'Debug.Print "GROUP"
                Err.Clear
                If .Fill.Visible = msoTrue Then 'Change Group color if needed
                    If CStr(.Fill.ForeColor.RGB) > 0 Then
                        If Err.Number = 0 Then
                            .Fill.ForeColor.RGB = getColorConversion(.Fill.ForeColor.RGB)
                        End If
                    End If
                End If
                Err.Clear

               For x = 1 To sld.Shapes(lShp).GroupItems.Count
                   Call changeShapeColor(sld.Shapes(lShp).GroupItems(x))
               Next x
            End If
        End If
        '--------------------------
      
      End With
Next
 
 
End Sub


Sub changeShapeColor(oSh)
    With oSh
        On Error GoTo skip1
        If .Fill.Visible = msoTrue Then
            If CStr(.Fill.ForeColor.RGB) <> 0 Then
                .Fill.ForeColor.RGB = getColorConversion(.Fill.ForeColor.RGB)
                .Fill.BackColor.RGB = getColorConversion(.Fill.BackColor.RGB)
            End If
        End If
        
skip1:
        On Error GoTo skip2
        If .Line.Visible = msoTrue Then
            If CStr(.Line.ForeColor.RGB) <> 0 Then
                .Line.ForeColor.RGB = getColorConversion(.Line.ForeColor.RGB)
            End If
        End If
        
skip2:
    
    End With

End Sub


Function getColorConversion(rgbColor, Optional darkMode)
Dim SourceRGBColor, DestRGBColor
Dim i As Long
Dim sourceBackgroundRGB1 As Long, sourceBackgroundRGB2 As Long, sourceAccentRGB1 As Long, sourceAccentRGB2 As Long, sourceAccentRGB3 As Long, sourceAccentRGB4 As Long, sourceAccentRGB5 As Long, sourceAccentRGB6 As Long
Dim sourceTextRGB1 As Long, sourceTextRGB2 As Long, sourceTextRGB3 As Long
Dim destBackgroundRGB1 As Long, destBackgroundRGB2 As Long, destAccentRGB1 As Long, destAccentRGB2 As Long, destAccentRGB3 As Long, destAccentRGB4 As Long, destAccentRGB5 As Long, destAccentRGB6 As Long
Dim destAccentRGB7 As Long, destAccentRGB8 As Long, destAccentRGB9 As Long, destAccentRGB10 As Long
Dim destTextRGB1 As Long, destTextRGB2 As Long, destTextRGB3 As Long, destTextRGB4 As Long


'----- source colors ----------
sourceBackgroundRGB1 = RGB(255, 255, 255)
sourceBackgroundRGB2 = RGB(15, 28, 80)
sourceAccentRGB1 = RGB(194, 206, 230)
sourceAccentRGB2 = RGB(103, 133, 193)
sourceAccentRGB3 = RGB(230, 182, 0)
sourceAccentRGB4 = RGB(188, 67, 40)
sourceAccentRGB5 = RGB(131, 178, 84)
sourceAccentRGB6 = RGB(170, 60, 128)
sourceTextRGB1 = RGB(64, 64, 64)
sourceTextRGB2 = RGB(0, 128, 177)
sourceTextRGB3 = RGB(0, 0, 0)

'-------------------------------

'----- alt colors ----------
destBackgroundRGB1 = RGB(7, 15, 38)
destBackgroundRGB2 = RGB(255, 255, 255)
destAccentRGB1 = RGB(0, 114, 188)
destAccentRGB2 = RGB(0, 91, 150)
destAccentRGB3 = RGB(25, 163, 252)
destAccentRGB4 = RGB(0, 203, 93)
destAccentRGB5 = RGB(0, 223, 237)
destAccentRGB6 = RGB(148, 148, 148)
destTextRGB1 = RGB(0, 0, 0)
destTextRGB2 = RGB(7, 15, 38)

destTextRGB3 = RGB(46, 64, 77)
destTextRGB4 = RGB(255, 255, 255)
destAccentRGB7 = RGB(255, 196, 0)
destAccentRGB8 = RGB(228, 38, 0)
destAccentRGB9 = RGB(255, 122, 0)
destAccentRGB10 = RGB(178, 32, 0)

'-----------------------------

SourceRGBColor = Array(sourceBackgroundRGB1, sourceBackgroundRGB2, sourceAccentRGB1, sourceAccentRGB2, sourceAccentRGB3, sourceAccentRGB4, sourceAccentRGB5, sourceAccentRGB6, sourceTextRGB1, sourceTextRGB2, sourceTextRGB3)

If IsMissing(darkMode) Then 'light mode
    DestRGBColor = Array(destBackgroundRGB2, destAccentRGB2, destAccentRGB1, destAccentRGB3, destAccentRGB7, destAccentRGB10, destAccentRGB4, destAccentRGB6, destTextRGB2, destTextRGB3, destTextRGB1)
Else 'dark mode
    DestRGBColor = Array(destBackgroundRGB2, destAccentRGB2, destAccentRGB1, destAccentRGB3, destAccentRGB7, destAccentRGB10, destAccentRGB4, destAccentRGB6, destTextRGB4, destTextRGB4, destTextRGB4)
End If

getColorConversion = rgbColor

For i = 0 To UBound(SourceRGBColor)
 If SourceRGBColor(i) = rgbColor Then
     getColorConversion = DestRGBColor(i)
     Exit Function
 End If
Next i

End Function


Private Sub deleteAdditionalPattern()
    Dim layoutNumber As Integer, i As Integer
    Const maxLayoutNumber = 22
    
    layoutNumber = ActivePresentation.Designs(1).SlideMaster.CustomLayouts.Count
    
    If layoutNumber > maxLayoutNumber Then
      For i = layoutNumber To maxLayoutNumber + 1 Step -1
         ActivePresentation.Designs(1).SlideMaster.CustomLayouts(i).Delete
      Next i
    End If
End Sub

Private Sub deleteAdditionalPattern2()
    Dim layoutNumber As Integer, i As Integer
    Const maxLayoutNumber = 4
    
    layoutNumber = ActivePresentation.Designs.Count
    
    If layoutNumber > maxLayoutNumber Then
      For i = layoutNumber To maxLayoutNumber + 1 Step -1
         ActivePresentation.Designs(i).Delete
      Next i
    End If
End Sub

Sub checkFooters()
  Dim oSld As Slide, i As Long, sourceSlideName As String
  
  
  On Error Resume Next
  

For i = 1 To ActivePresentation.Slides.Count


  '==== Update progress bar ===
    strProgress = i * 100 / ActivePresentation.Slides.Count
    ProgressBarUserForm.ProgressLabel.Caption = "Check Footers, " & Round(strProgress, 1) & "%, please wait...."
    ProgressBarUserForm.ProgressBar.Width = strProgress * 2
    ProgressBarUserForm.Show
    DoEvents
    '==============

  Set oSld = ActivePresentation.Slides.Item(i)
  
  sourceSlideName = LCase(ActivePresentation.Slides.Item(i).CustomLayout.Name)
  
  
   ' .....FOOTER.......

    With oSld.HeadersFooters
    
        .Footer.Visible = True
    
        .SlideNumber.Visible = True
    
        .DateAndTime.Visible = True
    
        .DateAndTime.UseFormat = True
    
        .DateAndTime.Format = ppDateTimeMdyy
    End With
    ''............

Next i

Unload ProgressBarUserForm

End Sub

Sub getShapePositionBack(sld As Slide, sourceSlideName As String, shapeArray)
Dim shapeFeatures

'------  get the shape back -------
On Error Resume Next

If sourceSlideName = "title" Or sourceSlideName = "sub title" Or sourceSlideName = "1_sub title" Then
   GoTo skipchangeshape
End If

changeshape:
If shapeArray(0) <> "" Then

    For j = 0 To UBound(shapeArray)
      shapeFeatures = Split(shapeArray(j), ";")
      If sld.Shapes(shapeFeatures(0)).Top <= 500 Then   'do not change footers
         
        If sld.Shapes(shapeFeatures(0)).Top < 20 And sld.Shapes(shapeFeatures(0)).Width > 850 Then GoTo nextshape  'do not change  titles
      
        sld.Shapes(shapeFeatures(0)).Top = shapeFeatures(1)
        sld.Shapes(shapeFeatures(0)).Left = shapeFeatures(2)
        
        On Error GoTo 0
        On Error Resume Next
        If sld.Shapes(shapeFeatures(0)).HasTextFrame = True And Len(sld.Shapes(shapeFeatures(0)).TextFrame.TextRange.Text) > 1 Then
          If Err.Number = 0 Then
            sld.Shapes(shapeFeatures(0)).Height = shapeFeatures(3)
            sld.Shapes(shapeFeatures(0)).Width = shapeFeatures(4)
          End If
        End If
      
      
      Else
      
        'sld.Shapes(shapeFeatures(0)).TextFrame.TextRange.Font.Color = RGB(255, 255, 255) 'set white
      
      End If
nextshape:
    Next j
End If
On Error GoTo 0
'----------------------------------

skipchangeshape:
End Sub


Function getLowestCustomName(sld As Slide, srtCustomName)
Dim shp As Shape
Dim strNumber
Dim auxString

getLowestCustomName = ""
auxString = ""

For Each shp In sld.Shapes
    If shp.HasTextFrame And InStr(shp.Name, srtCustomName) <> 0 Then
    
        strNumber = Trim(Replace(shp.Name, srtCustomName, ""))
        
        If IsNumeric(strNumber) Then
           If auxString = "" Then
             auxString = strNumber
           Else
             If strNumber < auxString Then
               auxString = strNumber
             End If
           End If
        
        End If
    
    End If

Next
getLowestCustomName = auxString
End Function

Private Sub changeBullets12Numbers(sld As Slide)
Dim i As Long

On Error Resume Next
    For lShp = sld.Shapes.Count To 1 Step -1
    
       With sld.Shapes(lShp)
    
               For i = 1 To .TextFrame.TextRange.Paragraphs.Count
                 If .TextFrame.TextRange.Paragraphs(i).ParagraphFormat.Bullet.Visible <> 0 Then
                   
                   With .TextFrame.TextRange.Paragraphs(i).ParagraphFormat.Bullet
    
                      .Type = 2
                   End With
                 End If
               Next i
    
      End With
    Next
End Sub

Private Sub deleteOnlyPicturesInSlide(sld As Slide)
    Dim shp As Shape
    For Each shp In sld.Shapes
       If Not shp.HasTextFrame Then
         shp.Delete
       End If
    Next
End Sub
