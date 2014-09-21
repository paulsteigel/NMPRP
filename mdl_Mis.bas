Option Explicit

Sub AssignCaption()
    Dim OLEObj As OLEObject
    Dim tSheet As Worksheet
    For Each tSheet In ThisWorkbook.Sheets
        For Each OLEObj In tSheet.OLEObjects
            If TypeOf OLEObj.Object Is MSForms.CheckBox Then
                OLEObj.Object.Caption = GetObjectCaption(tSheet.Name, OLEObj.Name)
            End If
        Next OLEObj
    Next tSheet
End Sub

Private Function GetObjectCaption(ParentName As String, ObjName As String) As String
    Dim rng As Range, Found As Boolean
    Set rng = Range("OBJ_CAPTION")
    While Not Found
        With rng
            If Len(rng) > 0 Then
                If .Value = ParentName And .Offset(0, 1) = ObjName Then
                    GetObjectCaption = .Offset(0, 2)
                    Found = True
                End If
            Else
                Found = True
            End If
        End With
        Set rng = rng.Offset(1)
    Wend
    Set rng = Nothing
End Function

Sub GetCaption()
    Dim rng As Range
    Set rng = Range("OBJ_CAPTION")
    Dim OLEObj As OLEObject
    Dim tSheet As Worksheet
    For Each tSheet In ThisWorkbook.Sheets
        For Each OLEObj In tSheet.OLEObjects
            If TypeOf OLEObj.Object Is MSForms.CheckBox Then
                rng = tSheet.Name
                rng.Offset(0, 1) = OLEObj.Name
                rng.Offset(0, 2) = OLEObj.Object.Caption
                rng.Offset(0, 3) = OLEObj.Object.Value
            End If
        Next OLEObj
    Next tSheet
End Sub

Sub SetProtection()
    Dim iSht As Worksheet, theCell As Range
    For Each iSht In ThisWorkbook.Worksheets
        If Len(iSht.Name) = 2 Then
            With iSht
                .Unprotect
                For Each theCell In .UsedRange
                    If Not theCell.Locked Then
                        With theCell
                            .Font.Bold = True
                            .HorizontalAlignment = xlLeft
                            .VerticalAlignment = xlTop
                            .WrapText = True
                            .IndentLevel = 1
                        End With
                    End If
                Next
                .Protect
            End With
        End If
    Next
End Sub

Sub SetPageWidth()
'
' SetPageWidth Macro
' Macro recorded 5/21/2013 by Dang Dinh Ngoc
'

'
    Dim theSheet As Worksheet
    For Each theSheet In ThisWorkbook.Sheets
        With theSheet.PageSetup
            '.Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .PrintErrors = 0
            '.LeftHeader = ""
            '.CenterHeader = ""
            '.RightHeader = ""
            '.LeftFooter = ""
            .CenterFooter = "Trang: &P/&N"
            '.RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.4)
            .RightMargin = Application.InchesToPoints(0.4)
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.6)
            .HeaderMargin = Application.InchesToPoints(0.25)
            .FooterMargin = Application.InchesToPoints(0.25)
            '.PrintHeadings = False
            '.PrintGridlines = False
            '.PrintComments = xlPrintNoComments
            '.PrintQuality = 300
            .CenterHorizontally = True
            '.CenterVertically = False
            .Orientation = xlPortrait
            .Draft = False
            .PaperSize = xlPaperA4
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = 100
            '.PrintErrors = xlPrintErrorsDisplayed
        End With
    Next
End Sub

Sub Tes1()
' Apply changes made to relative button if there may have some...
    'If ActiveSheet.Name <> "Data" Then Exit Sub
    Dim ShObj As Shape, Obj As Object
    Dim i As Long
    
    'XUnProtectSheet Sheet5
    Set ShObj = Sheet5.Shapes("Group 60").GroupItems(1)
    For Each ShObj In Sheet5.Shapes
        Debug.Print ShObj.Name
        
        If ShObj.Type = msoGroup Then
            For Each Obj In ShObj.GroupItems
                Debug.Print Obj.Name
                
            Next
        End If
    Next
    'For i = 0 To ShObj.GroupItems.Count
    '    Debug.Print ShObj.GroupItems(i).Name
    'Next
    ' Just block it now and save
    'ShObj.TextFrame.Characters.Text = CStr(Range("LST_TRAIN_TYPE").Cells(1))
    'Debug.Print ShObj.Name
    'XProtectSheet Sheet5
End Sub

Sub TestAccess()
    GetControl Sheet5
End Sub

Private Sub GetControl(SheetObj As Worksheet, Optional CtrObj As Shape)
    Dim Obj As Shape
    If CtrObj Is Nothing Then
        For Each Obj In SheetObj.Shapes
            GetControl SheetObj, Obj
        Next
    Else
        If CtrObj.Type = msoGroup Then
            For Each Obj In CtrObj.GroupItems
                If Obj.Type = msoGroup Then
                    GetControl SheetObj, Obj
                Else
                    On Error Resume Next
                    Debug.Print Obj.Name & "/" & Obj.TextFrame.Characters.Text
                    'Obj.DrawingObject.Characters.Text = "xxx"
                End If
            Next
        End If
    End If
End Sub

Function WhichOption(shpGroupBox As Shape) As OptionButton
    Dim shp As OptionButton
    Dim shpOptionGB As GroupBox
    Dim gb As GroupBox

    If shpGroupBox.Type <> msoGroup Then Exit Function
    Set gb = shpGroupBox.OLEFormat.Object
    For Each shp In shpGroupBox.Parent.OptionButtons
        Set shpOptionGB = shp.GroupBox
        If Not shpOptionGB Is Nothing Then
            If shpOptionGB.Name = gb.Name Then
                If shp.Value = 1 Then
                    Set WhichOption = shp
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Sub testXX()
    Dim shpOpt As OptionButton

    Set shpOpt = WhichOption(Worksheets("T5").Shapes("Group 60"))
    Debug.Print shpOpt.Name
End Sub

