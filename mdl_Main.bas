Option Explicit

' For storing item attribute
Private Type ItemAttributes
    ItemDetails As String
    ItemHeading As String
    ItemEmphasizeHeading As String
    DataSource As String
    Name As String
End Type

Public Type TextObject
    TextCount As Long
    TextValue As String
End Type

Public Type ObjectEquation
    VariableName As String
    VariableFomular As String
End Type

Dim OldSheet As Boolean

' Cached variable for keeping some temporary stuff
Private CachedListDistinct As Collection
Private ColListing() As New Collection
Private CurrentPointer As Long
Private OldTableName As String

Sub Back2Main()
    ' For returning to Main Screen
    ActivateSheet "Manhinhchinh"
End Sub

Private Sub ActivateSheet(SheetName As String)
    ThisWorkbook.Sheets(SheetName).Activate
End Sub

Sub Act_II_2_A()
    If OldSheet Then
        ActivateSheet "II.2"
    Else
        ActivateSheet "II.2.A"
    End If
    OldSheet = Not OldSheet
End Sub

Sub Act_II_2_B()
    ActivateSheet "II.2.B"
End Sub

Sub Act_II_5_A()
    ActivateSheet "II.5.A"
End Sub

Sub Act_II_5_B()
    ActivateSheet "II.5.B"
End Sub

Sub Act_II_6_E()
    If Range("CONF_SCORE") <> 1 Then Exit Sub
    ActivateSheet "II.6.E"
    Range("COND_FLOOR").Activate
End Sub

Sub CriteriaEditor()
    ' Activate form for creating criteria
End Sub

Sub EvaluateActivity()
    ' Show form to conduct evaluation...
End Sub

Function CreateWordDocument(retApp As Object) As Object
    'Muc dich: Co gang thiet lap ket noi voi mot phien lam viec cua Word neu duoc. Neu khong thi tao moi
    Dim wrdApp As Object
    
    'Co gang tao ket noi
    On Error Resume Next
    Set wrdApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        'Khong tao duoc ketnoi
        Err.Clear
        Set wrdApp = CreateObject("Word.Application")
        wrdApp.Visible = True
    End If
    ' Doan code chinh...
    Set retApp = wrdApp
    Set CreateWordDocument = wrdApp.Documents.Add
End Function

Sub GenerateSEDP()
    RegisterAction
    Application.StatusBar = ""
    ' First - convert all to Unicode
    AppInit
    
    ' Turn off some stuff...
    ShowOff
    
    SheetObjName = "II.5.A"
    ConvertRange Range("tblUnicode_1")
    
    SheetObjName = "II.5.B"
    ConvertRange Range("tblUnicode_2")
    
    'reset some collections
    Set CachedListDistinct = Nothing
    ReDim ColListing(0)
    
    Dim myWordDoc As Object, LocalSetting As String
    LocalSetting = ","
    If InStr(Format("12345", "#,##0"), ",") > 0 Then LocalSetting = "."
       
    Dim myWordApp As Object
    Set myWordDoc = CreateWordDocument(myWordApp)
    
    Dim i As Long, ContractDoc As String, HasWordError As Boolean, IntegrationSetting As Boolean
    Dim MyRange As Range
    myWordApp.Visible = True
    Set MyRange = Range("SEDP_OUTLINE")
    
    ' now generate all style
    HasWordError = GenerateWordStyle(myWordDoc, myWordApp)
    If HasWordError Then GoTo errHandler
    Dim j, k, L, x As Long, xCounter As Long, prCount As Long
    
    Dim FilterStr As String, FilterArr As Variant, StructStr As String, AllRowCount As Long
    
    ' Storing all atribute of the items
    Dim StyleLevel1 As ItemAttributes, StyleLevel2 As ItemAttributes, StyleItems() As ItemAttributes, StyleSpecs() As ItemAttributes
    Dim tmpObj As ItemAttributes, ResetObj As ItemAttributes, ResetObjArr() As ItemAttributes
    Dim ObjOpt() As ObjectEquation, xPos As Long
    
    Dim MsgPasstoWord As String, MsgPassSectortoWord As String, MsgGetToNewSystem As String, MsgFinished As String
    MsgPasstoWord = MSG("MSG_PASS_TO_WORD")
    MsgPassSectortoWord = MSG("MSG_PASS_SECTOR_TO_WORD")
    MsgGetToNewSystem = MSG("MSG_GET_TO_NEW_SYSTEM")
    MsgFinished = MSG("MSG_FINISHED")

    ' Now conduct a sort first
    SortTable ThisWorkbook, "II.5.A", "tblUnicode_1", "A6", "B6"
    ' will force this sorting later
    SortTable ThisWorkbook, "II.5.C", "tblUnicode_3", "C6", "A6"
    SortTable ThisWorkbook, "II.5.D", "tblUnicode_4", "B6", "A6"
    ' Prepare all Sheets for printing out
    ApplySheetFilter
    
    IntegrationSetting = Range("CONF_INTEGRATE")
    With myWordDoc
        i = 2
        AllRowCount = MyRange.Rows.Count
        While i <= AllRowCount
            If Left(MyRange.Cells(i, 1), 6) = "[DATA]" Then
                ' Start processing the internal addon
                For j = 1 To Range("tblKeySector").Rows.Count
                    If StyleLevel1.ItemHeading = "" Then
                        StyleLevel1.ItemHeading = MyRange.Cells(i, 5)
                        StyleLevel1.ItemDetails = MyRange.Cells(i, 2)
                        StyleLevel1.ItemEmphasizeHeading = MyRange.Cells(i, 4)
                    End If
                    ' Now add data
                    InsertPara myWordDoc, StyleLevel1, Range("tblKeySector").Cells(j, 2)
                    
                    ' Okie - move to next level
                    While CStr(Left(Range("tblKeySubSector").Cells(k + 1, 1), 1)) = CStr(Range("tblKeySector").Cells(j, 1))
                        If StyleLevel2.ItemHeading = "" Then
                            StyleLevel2.ItemHeading = MyRange.Cells(i + 1, 5)
                            StyleLevel2.ItemDetails = MyRange.Cells(i + 1, 2)
                            StyleLevel2.ItemEmphasizeHeading = MyRange.Cells(i + 1, 4)
                        End If
                        ' Now add data
                        InsertPara myWordDoc, StyleLevel2, Range("tblKeySubSector").Cells(k + 1, 2)
                    
                        If xCounter > 0 Then GoTo xProcess
                        ' Caching style object
                        i = i + 1
                        While MyRange.Cells(i, 3) <> ""
                            If MyRange.Cells(i, 3) = "I" Then
                                ReDim Preserve StyleItems(xCounter)
                                StyleItems(xCounter).ItemHeading = MyRange.Cells(i, 5)
                                StyleItems(xCounter).ItemDetails = MyRange.Cells(i, 2)
                                StyleItems(xCounter).ItemEmphasizeHeading = MyRange.Cells(i, 4)
                                StyleItems(xCounter).Name = MyRange.Cells(i, 1)
                            ElseIf MyRange.Cells(i, 3) = "S" Then
                                ReDim Preserve StyleSpecs(xCounter)
                                StyleSpecs(xCounter).ItemHeading = MyRange.Cells(i, 5)
                                StyleSpecs(xCounter).ItemDetails = MyRange.Cells(i, 2)
                                StyleSpecs(xCounter).ItemEmphasizeHeading = MyRange.Cells(i, 4)
                                StyleSpecs(xCounter).DataSource = MyRange.Cells(i, 4)
                                StyleSpecs(xCounter).Name = MyRange.Cells(i, 1)
                                xCounter = xCounter + 1
                            End If
                            i = i + 1
                        Wend
                        ' By ending of this line, I already stop at the next processing row
xProcess:               ' Get data on the fly
                        ' Now we have to build the list of criteria
                        While CStr(Range("tblKeySubSectorItems").Cells(L + 1, 2)) = CStr(Range("tblKeySector").Cells(j, 1) & "." & Replace(CStr(Range("tblKeySubSector").Cells(k + 1, 1)), LocalSetting, "."))
                            FilterStr = FilterStr & "/" & Range("tblKeySubSectorItems").Cells(L + 1, 3)
                            L = L + 1
                        Wend
                        ' Now add data
                        'Debug.Print FilterStr
                        For x = 0 To xCounter - 1
                            ' add heading
                            ' we also need to deal with this stuff...
                            InsertPara myWordDoc, StyleItems(x), IIf(StyleItems(x).Name <> "", StyleItems(x).Name & " ", "") & Replace(StyleItems(x).ItemDetails, "[RELITEM]", Range("tblKeySubSector").Cells(k + 1, 2)), True
                            ' add details for sub-items
                            InsertPara myWordDoc, StyleSpecs(x), GetFilteredData(FilterStr, StyleSpecs(x).DataSource)
                        Next
                        k = k + 1 ' increase second level text
                        FilterStr = ""
                        ' throw status on sectorising
                        Application.StatusBar = MsgPassSectortoWord & " " & Format(k * 100 / 13, "##0") & "% " & MsgFinished
                    Wend
                Next
                ' reset all variables and styles
                StyleLevel1 = ResetObj
                StyleLevel2 = ResetObj
                StyleItems() = ResetObjArr()
                StyleSpecs() = ResetObjArr()
                xCounter = 0
                k = 0
                L = 0
                ' get back one row
                i = i - 1
            Else
                ' check for intgration
                If Not IntegrationSetting And MyRange.Cells(i, 3) = "x" Then GoTo SKIP_INTEGRATION
                With tmpObj
                    .ItemHeading = MyRange.Cells(i, 5)
                    .ItemDetails = MyRange.Cells(i, 2)
                    .ItemEmphasizeHeading = "Normal"
                End With
                
                StructStr = MyRange.Cells(i, 4)
                If Left(StructStr, 8) = "[[:TABLE" Then
                    ' First insert normal text
                    InsertPara myWordDoc, tmpObj, MyRange.Cells(i, 1)
                    ' Now insert the table
                    InsertTable myWordDoc, MyRange.Cells(i, 4)
                ElseIf StructStr = "INCLUDED" Then
                    If InStr(MyRange.Cells(i, 2), "tblUnicode") <> 0 Then
                        'Now retrieve some counting function first
                        ObjOpt = GetOption(MyRange.Cells(i, 2))
                        StructStr = MyRange.Cells(i, 1)
                        For xPos = LBound(ObjOpt) To UBound(ObjOpt)
                            StructStr = Replace(StructStr, ObjOpt(xPos).VariableName, ObjOpt(xPos).VariableFomular)
                        Next
                        InsertPara myWordDoc, tmpObj, StructStr, True
                    Else
                        ' Just sometinh else
                        Debug.Print "xx"
                    End If
                ElseIf StructStr = "REPEAT" Then
                    ' okie they need to repead this stuff..
                    ' First -see howmany repeat stuff...
                    j = i ' remember the first row started
                    
                    For CurrentPointer = 0 To UBound(ColListing)
                        i = j
                        While MyRange.Cells(i, 4) = "REPEAT"
                            With tmpObj
                                .ItemHeading = MyRange.Cells(i, 5)
                                .ItemDetails = MyRange.Cells(i, 2)
                                .ItemEmphasizeHeading = "Normal"
                            End With
                            'Now retrieve some counting function first
                            If Trim(MyRange.Cells(i, 2)) <> "" Then
                                ObjOpt = GetOption(MyRange.Cells(i, 2))
                                
                                StructStr = MyRange.Cells(i, 1)
                                If CurrentPointer > 0 Then
                                    ' remove top comments
                                    StructStr = Replace(StructStr, "[COMMENT_TOP_EVENT]", "")
                                End If
                                For xPos = LBound(ObjOpt) To UBound(ObjOpt)
                                    StructStr = Replace(StructStr, ObjOpt(xPos).VariableName, ObjOpt(xPos).VariableFomular)
                                Next
                                InsertPara myWordDoc, tmpObj, StructStr, True
                            Else
                                InsertPara myWordDoc, tmpObj, MyRange.Cells(i, 1), True
                            End If
                            i = i + 1
                        Wend
                    Next
                    'get back i a step
                    'reset this variable
                    CurrentPointer = 0
                    i = i - 1
                Else
                    ' Just insert the normal text
                    InsertPara myWordDoc, tmpObj, MyRange.Cells(i, 1)
                End If
            End If
SKIP_INTEGRATION:
            i = i + 1
            Application.StatusBar = MsgPasstoWord & " " & Format((i - 2) * 100 / AllRowCount, "##0") & "% " & MsgFinished
        Wend
        ReformatWordTable myWordDoc
    End With

    ' get out and close
    SaveFile ThisWorkbook.Path & "\" & Range("SEDP_Name") & Format(Now(), "HHMM_DDMMYYYY") & ".doc", myWordDoc
    
    Application.StatusBar = MSG("MSG_DONE_ALL")
    
errHandler:
    If HasWordError Then
        Err.Clear
        MsgBox MSG("MSG_WORD_NOT_CLOSE"), vbCritical
    'Else
    '    MsgBox MSG("MSG_FINISHED_CREATE_PLAN"), vbInformation + vbOKOnly
    End If
    myWordApp.Visible = True
    myWordApp.Activate
    Set myWordDoc = Nothing
    Set myWordApp = Nothing
    
    ShowOff True

    DeRegisterAction
End Sub

Private Sub SaveFile(FileName, DocObj As Object)
    On Error GoTo errHandler
    DocObj.Paragraphs(1).Range.Delete
    If Dir(FileName) <> "" Then Kill FileName
    DocObj.SaveAs FileName
errHandler:
    If Err.Number <> 0 Then
        MsgBox MSG("MSG_SAVE_FALSE"), vbCritical
    End If
End Sub

Private Sub InsertPara(DocObj As Object, ItemStyle As ItemAttributes, ItemText As String, Optional OverideAdd As Boolean = False)
    'On Error Resume Next
    Dim prCount As Long, tmpText As String, tmpItem As ItemAttributes
    tmpItem = ItemStyle
    With DocObj
        If ItemStyle.ItemHeading = "" Or ItemText = "" Then Exit Sub
        .Paragraphs.Add
        prCount = .Paragraphs.Count
        .Paragraphs(prCount).Range.Style = .Styles(ItemStyle.ItemHeading)
        .Paragraphs(prCount).Range.Text = ItemText
        
        If ItemStyle.ItemDetails <> "" And Not OverideAdd Then
            ' Add new introduction line if neccessary
            tmpItem.ItemHeading = tmpItem.ItemEmphasizeHeading
            tmpText = tmpItem.ItemDetails
            tmpItem.ItemDetails = ""
            InsertPara DocObj, tmpItem, tmpText
        End If
    End With
End Sub

Private Sub InsertTable(DocObj As Object, RangeName As String)
    Dim prCount As Long, tmpObj As Object, CopyRange As Range
    Dim RngName As Variant, ColIndex As Variant
    Dim tmpWbk As Workbook, tmpSheet As Worksheet, i As Long
    Dim FilterColumn As Long, FilterObject As String
    Dim UseHeader As Boolean ' sometimes forgot to get the header of the table
    Dim Row2Copy As Long
    
    ' For inputdata
    RngName = Split(RangeName, "/")
    ' For showing column
    ColIndex = Split(RngName(2), ",")
    ' For column to limit
    FilterColumn = RngName(3)
    If RngName(4) <> "" Then FilterObject = Evaluate(RngName(4))
    UseHeader = Evaluate(RngName(5))
    ' Assign Range now
    Set CopyRange = Range(RngName(1))
    ' Now create a new workbook and format the table
    Set tmpWbk = Workbooks.Add
    Set tmpSheet = tmpWbk.Sheets.Add
    If UseHeader Then Set CopyRange = CopyRange.Resize(CopyRange.Rows.Count + 1).Offset(-1)
    CopyRange.Copy
    tmpSheet.Range("B1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    
    ' Now change column size
    For i = 1 To CopyRange.Columns.Count
        tmpSheet.Columns(i + 1).ColumnWidth = CopyRange.Columns(i).ColumnWidth ' ThisWorkbook.Sheets("II.2.B").Columns(i).Width
    Next
    ' Now remove some rows if needed
    If FilterColumn > 0 Then
        Dim tCell As Range, DeletedAlready As Boolean
        
        Set tCell = tmpSheet.Cells(1, FilterColumn + 1)
        While tCell <> ""
            Row2Copy = tCell.Row
            If FilterObject <> "" Then
                If tCell = FilterObject Then
                    tCell.EntireRow.Delete
                    DeletedAlready = True
                End If
            End If
            
            If DeletedAlready Then
                Set tCell = tmpSheet.Cells(Row2Copy, FilterColumn + 1)
            Else
                Set tCell = tCell.Offset(1)
            End If
            DeletedAlready = False
            Row2Copy = tCell.Row
        Wend
        Row2Copy = Row2Copy - 1
    Else
        Row2Copy = CopyRange.Rows.Count
    End If
    ' Now disable some columns
    ' Build a string with column to be removed
    ' Remove some blank line
    
    ' Continue the next
    Dim tmpStr As String, relCol As Variant
    For i = 4 To CopyRange.Columns.Count
        tmpStr = tmpStr & "," & i
    Next
    
    For i = UBound(ColIndex) To LBound(ColIndex) Step -1
        If Val(ColIndex(i)) > 3 Then
            tmpStr = Replace(tmpStr, "," & CStr(ColIndex(i)), "")
        Else
            Exit For
        End If
    Next
    relCol = Split(tmpStr, ",")
    For i = UBound(relCol) To LBound(relCol) Step -1
        If Val(relCol(i)) > 3 Then
            tmpSheet.Columns(Val(relCol(i)) + 1).Delete Shift:=xlToLeft
        Else
            Exit For
        End If
    Next
    ' Now just copy them to word
    Set CopyRange = tmpSheet.Range("B1", tmpSheet.Cells(Row2Copy, UBound(ColIndex) + 2))
    With DocObj
        CopyRange.Copy
        prCount = .Paragraphs.Count
        .Paragraphs(prCount).Range.PasteExcelTable False, True, True
        Set tmpObj = .Tables(.Tables.Count)
        With tmpObj
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 100
            .Rows.HeightRule = wdRowHeightAtLeast
            .Rows.Height = Excel.Application.CentimetersToPoints(0)
            .Rows.LeftIndent = Excel.Application.CentimetersToPoints(0)
        End With
    End With
    Application.CutCopyMode = False
    Set tmpObj = Nothing
    Set tmpSheet = Nothing
    tmpWbk.Close False
    Set tmpWbk = Nothing
End Sub

Private Sub ReformatWordTable(WrdDoc As Object)
    Dim tmpObj As Object, i As Long
    For Each tmpObj In WrdDoc.Tables
        'Format the selected table
        With tmpObj.Range.ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = wdLineSpaceSingle
            .FirstLineIndent = Excel.Application.CentimetersToPoints(0)
        End With
        With tmpObj
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 100
            .Rows.HeightRule = wdRowHeightAtLeast
            .Rows.Height = Excel.Application.CentimetersToPoints(0)
            .Rows.LeftIndent = Excel.Application.CentimetersToPoints(0)
        End With
        ' Remove trailing space
        For i = 1 To 10
            With tmpObj.Range.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "  "
                .Replacement.Text = " "
                .Forward = True
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
        Next
    Next
    Set tmpObj = Nothing
End Sub

Private Function CountTable(Obj As Object) As Long
    On Error GoTo errHandler
    CountTable = Obj.Tables.Count
errHandler:
End Function

Private Function GetFilteredData(iFilter As String, iColumn As String) As String
    'Base on the defined filter, try to get somedata from this - Don't care data range for
    Dim SrcArr As Variant, SrcRange As Range, i As Long, lRetStr  As String
    Dim OldText As String
    
    If Val(iColumn) <= 0 Then Exit Function
    Set SrcRange = Range("tblUnicode_1")
    i = 1
    While SrcRange.Cells(i, 1) <> ""
        If InStr(iFilter & "/", "/" & SrcRange.Cells(i, 1) & "/") <> 0 Then
            ' I found first ocurrence of the text
            OldText = SrcRange.Cells(i, 1)
            While SrcRange.Cells(i, 1) = OldText
                lRetStr = lRetStr & "//" & SrcRange.Cells(i, Val(iColumn))
                i = i + 1
            Wend
        Else
            i = i + 1
        End If
    Wend
    If lRetStr <> "" Then
        lRetStr = Replace(Mid(lRetStr, 3), "//", vbLf)
        GetFilteredData = lRetStr
    End If
End Function

Sub SortTable(WrbObj As Workbook, WksObjName As String, RngName As String, SortKey1 As String, Optional SortKey2 As String)
    ' This procedure will sort the selected table using sortkey
    Dim TheSheet As Worksheet
    Set TheSheet = WrbObj.Sheets(WksObjName)
    ' unprotect the sheet first
    XUnProtectSheet TheSheet
    'Activate the sheet
    WrbObj.Worksheets(WksObjName).Activate
    WrbObj.Worksheets(WksObjName).Range(RngName).Sort Key1:=Range(SortKey1), Order1:=xlAscending, Key2:=Range(SortKey2) _
        , Order2:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal
    
    ' ReProtect the sheet
    XProtectSheet TheSheet
    Set TheSheet = Nothing
End Sub

Private Sub AppInit()
    ' Now set all global variable
    CodeDestination = "Unicode"
    IsUpperText = False
    IsLowerText = False
    AutoCodeDetect = True
    ' Get code list
    CodeArray = SupportCodes
End Sub

Sub ActivateData()
    Sheets("Data").Activate
End Sub

Sub ActivateMain()
    Sheets("Main").Activate
End Sub

Sub UpdateII2B()
    'Update this sheet
    XUnProtectSheet Sheets("II.2.B")
    Dim TheRange As Range, CellFirst As Range, CellLast As Range
    Set CellFirst = Sheets("II.2.B").Range("II2BFIRST").Offset(1)
    Set CellLast = Sheets("II.2.B").Range("II2BLAST").Offset(-1)
    Set TheRange = Sheets("II.2.B").Range(CellFirst, CellLast)
    Set CellLast = CellLast.Offset(1, 5)
    'Copy formular
    CellLast.Copy
    TheRange.Offset(, 5).PasteSpecial xlPasteFormulas
    Application.CutCopyMode = False
    Set CellLast = CellLast.Offset(, 2)
    CellLast.Copy
    TheRange.Offset(, 7).PasteSpecial xlPasteFormulas
    Application.CutCopyMode = False
    Set CellLast = CellLast.Offset(, 2)
    CellLast.Copy
    TheRange.Offset(, 9).PasteSpecial xlPasteFormulas
    Application.CutCopyMode = False
    XProtectSheet Sheets("II.2.B")
End Sub

Private Sub Repair_II5A(Optional SheetName As String)
    ' Unprotect sheets
    XUnProtectSheet ThisWorkbook.Sheets(SheetName)
    With ThisWorkbook.Sheets(SheetName)
        .Range("A385:G385").Copy
        .Range("A6:G384").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone
        .Range("A6:G384").Locked = False
        .Activate
        .Range("A6").Select
    End With
    Application.CutCopyMode = False
    ' Reprotect sheet
    XProtectSheet ThisWorkbook.Sheets(SheetName)
End Sub

Private Sub Repair_II5B()
    ' Repair II.5.B
    XUnProtectSheet ThisWorkbook.Sheets("II.5.B")
    With ThisWorkbook.Sheets("II.5.B")
        .Range("J7").FormulaR1C1 = "=SUM(RC[1]:RC[4])"
        .Range("J7").Copy
        ' paste formular
        .Range("tblDataSumCol").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone
        Application.CutCopyMode = False
        .Range("B556:S556").Copy
        ' paste format
        .Range("tblUnicode_2").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone
        Application.CutCopyMode = False
        .Range("tblUnicode_2").PasteSpecial xlPasteValidation, Operation:=xlNone
        Application.CutCopyMode = False
        
        ' now paste validation
        'Dim VldRangeSrc As Range, VldRangeDst As Range, i As Long
        'Set VldRangeSrc = .Range("S556")
        'Set VldRangeDst = .Range("S7:S555")
        'For i = 1 To 14
        '    If i <= 11 Then
        '        VldRangeSrc.Copy
        '        VldRangeDst.PasteSpecial xlPasteValidation, xlPasteSpecialOperationNone
        '    Else
        '        VldRangeDst.Validation.Delete
        '    End If
        '    Set VldRangeSrc = VldRangeSrc.Offset(, -1)
        '    Set VldRangeDst = VldRangeDst.Offset(, -1)
        'Next
        .Activate
        .Range("C7").Select
    End With
    ' Unlock some areas
    ThisWorkbook.Sheets("II.5.B").Range("tblUnicode_2").Locked = False
    ' Reprotect the sheet
    XProtectSheet ThisWorkbook.Sheets("II.5.B")
End Sub

Sub RepairSheet(Optional SheetObj As String = "")
    ' This procedure shall repare all sheet.
    ShowOff
    If SheetObj = "" Then
        Repair_II5A "II.5.A"
        Repair_II5A "II.5.C"
        Repair_II5B
        ' update II.2.B
        UpdateII2B
    Else
        Select Case SheetObj
        Case "II.5.A", "II.5.C":
            Repair_II5A SheetObj
        Case "II.5.B":
            Repair_II5B
        Case "II.2.B":
            UpdateII2B
        Case Else
        End Select
    End If
    'Sheet11.Activate
    ShowOff True
End Sub

Sub ApplySheetFilter()
    'Activate filter on selected sheets
    ApplyFilter ThisWorkbook.Sheets("II.6.A"), "A7", 3, "<>"
    ApplyFilter ThisWorkbook.Sheets("II.6.B"), "A7", 3, "<>"
    ApplyFilter ThisWorkbook.Sheets("II.6.C"), "B5", 1, "<>"
    ApplyFilter ThisWorkbook.Sheets("II.2.B"), "I4", 1, "Có"
End Sub

Private Sub ApplyFilter(SheetObj As Worksheet, AppliedRange As String, FieldNum As Long, Criteria1 As String)
    XUnProtectSheet SheetObj
    SheetObj.Range(AppliedRange).AutoFilter Field:=FieldNum, Criteria1:=Criteria1
    XProtectSheet SheetObj
End Sub

Sub QuickFilter()
    Dim FldCriteria As String, FldNum As Long
    XUnProtectSheet ActiveSheet
    FldCriteria = "<>"
    Select Case ActiveSheet.Name
    Case "II.2.A":
        FldNum = 1
        FldCriteria = "Có"
    Case "II.2.B":
        FldNum = 1
        FldCriteria = "Có"
    Case "II.6.A", "II.6.B":
        FldNum = 3
    Case "II.6.C", "II.6.D", "II.6.F":
        FldNum = 1
    Case "II.5.A", "II.5.C", "II.5.D":
        FldNum = 1
    Case "II.5.B":
        ' Filter just nonacceptable stuff
        FldNum = 19
        ActiveSheet.Range(ActiveSheet.Name & "!_FilterDatabase").AutoFilter Field:=FldNum, _
            Criteria1:="=" & MSG("MSG_ST_NOTOK"), Operator:=xlOr, Criteria2:="=" & MSG("MSG_ST_VERIFY")
        GoTo ExitSub
    Case Else
        GoTo ExitSub
    End Select
    ActiveSheet.Range(ActiveSheet.Name & "!_FilterDatabase").AutoFilter Field:=FldNum, Criteria1:=FldCriteria
ExitSub:
    XProtectSheet ActiveSheet
End Sub

Sub ReleaseSheetFilter()
    ShowAll ThisWorkbook.Sheets("II.5.A")
    ShowAll ThisWorkbook.Sheets("II.5.B")
    ShowAll ThisWorkbook.Sheets("II.6.A")
    ShowAll ThisWorkbook.Sheets("II.6.B")
    ShowAll ThisWorkbook.Sheets("II.6.D")
    ShowAll ThisWorkbook.Sheets("II.6.C")
    ShowAll ThisWorkbook.Sheets("II.2.B")
End Sub

Sub XProtectSheet(s As Worksheet)
    If s.Name = "II.2.B" Then
        s.Protect "d1ndh1sk", Contents:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, DrawingObjects:=True, Scenarios:=True, _
        AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, UserInterfaceOnly:=True, AllowInsertingRows:=True, AllowDeletingRows:=True
    Else
        s.Protect "d1ndh1sk", Contents:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, DrawingObjects:=True, Scenarios:=True, _
        AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, UserInterfaceOnly:=True
    End If
End Sub

Sub XUnProtectSheet(s As Worksheet)
    s.Unprotect "d1ndh1sk"
End Sub

Sub ShowOff(Optional TurnEventOn As Boolean = False)
    ' Turn off everything, toggle
    Application.StatusBar = ""
    Application.ScreenUpdating = TurnEventOn
    Application.EnableEvents = TurnEventOn
    Application.CutCopyMode = False
    If TurnEventOn Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
End Sub

Sub MergeData(Optional SuccessFullCall As Boolean = False)
    ' This procedure shall help merging data from various table into this.
    ' By doing this, the application shall ask user from verifying some key question to make sure that they will not
    ' try to duplicate the import
    MsgBox MSG("MSG_IMPORT_LIMITED"), vbInformation
    'MSG_IMPORT_DISABLE

    Exit Sub
    '-----------------------------------------------------------------------
    
    ShowOff
    ' First - convert all to Unicode
    Dim SrcBook As Workbook, SrcSheet As Worksheet
    Dim DstBook As Workbook, DstSheet As Worksheet
    Dim ObjDlg As Dialog
    ' Now open the existing workbook to import data
    Set ObjDlg = Application.Dialogs.Item(xlDialogOpen)
    
    Dim FileLocation As String, FldBrowser As String
StepRetry:
    FldBrowser = GetFolder(ThisWorkbook.Path, True, "*.xls")
    If Dir(FldBrowser) = "" Then
        ' Exit code
        If MsgBox(MSG("MSG_SELECT_NO_FILE"), vbInformation + vbYesNo) = vbYes Then GoTo StepRetry
        GoTo StepEnd
    End If
    If FldBrowser = ThisWorkbook.Path & "\" & ThisWorkbook.Name Then
        MsgBox MSG("MSG_ERROR_THIS_FILE"), vbInformation
        GoTo StepRetry
    End If
    
    On Error GoTo StepEnd
    If FldBrowser = "" Then GoTo StepExit
    Set SrcBook = GetOpenWorkbook(FldBrowser)
    
    Set DstBook = ThisWorkbook
    ' check if this file is good or not
    If Not SheetValid("Data", SrcBook) Or Not SheetValid("II.5.A", SrcBook) Or Not SheetValid("II.5.B", SrcBook) Then
        MsgBox MSG("MSG_BAD_FILE"), vbCritical
        GoTo StepEnd
    End If
    ' Declaration of variables
    ' Initiate conversion procedures
    AppInit
    
    With SrcBook
        ' Ok - now, start moving data out a bit
        ' Data from sheet main
        Application.StatusBar = MSG("MSG_COPY_CONFIG")
        '////////////////////////////
        ' Revised on Apr 14 2013 with code name for wb project
        '////////////////////////////
        ParseRange SrcBook, DstBook, "Main", "FIG_CMN_NAME,FIG_PLN_YEAR,FIG_CUR_YEAR,FIG_CMN_CHAIR,FIG_CMN_ACCT,FIG_PLN_DATE,FIG_CMN_SIGNER,FIG_PLNST_DATE"
        ParseRange SrcBook, DstBook, "Data", "PrvCode,DstCode,CmnCode", True
                
        ' Get data from II.5.A & II.5.B - tbUnicode_1, 2
        ' Find out how many existing rows base on the activity column
        Application.StatusBar = MSG("MSG_COPY_II5A")
        CopyObject .Sheets("II.5.A"), SrcBook, "tblUnicode_1", DstBook.Sheets("II.5.A"), "tblUnicode_1", True
        
        
        ' Since user may use data from previous version.... we will have to consider the problem
        Application.StatusBar = MSG("MSG_COPY_II5B")
        CopyObject .Sheets("II.5.B"), SrcBook, "tblUnicode_2", DstBook.Sheets("II.5.B"), "tblUnicode_2"
        
        ' Now we should copy all Targeted program from opened table to existing one with little tricky stuff.
        Dim i As Long, prcCell As Range, prcString As String, dstCell As Range, curCnt As Long
        For Each prcCell In DstBook.Sheets("Data").Range("COND_TARGET").Cells
            prcString = "[" & prcCell & "]" & prcString
        Next
        ' now loop through the srcrange to add more stuff
        i = FindColHeader(.Sheets("Data"), 1, MSG("MSG_LIST_PROG"))
        Set prcCell = .Sheets("Data").Cells(8, i)
        While Len(Trim(prcCell)) > 0
            If InStr(prcString, "[" & prcCell & "]") <= 0 Then
                ' Make this value to existing sheet
                DstBook.Sheets("Data").Range("COND_TARGET").Cells(DstBook.Sheets("Data").Range("COND_TARGET").Rows.Count).Offset(1) = prcCell
            End If
            Set prcCell = prcCell.Offset(1)
        Wend
        ' now we copy village accross the form
        i = i - 6
        
        Set prcCell = .Sheets("Data").Cells(2, i)
        Set dstCell = DstBook.Sheets("Data").Range("tblVillageStart")
        
        ' Get numbers of current village columns
        curCnt = DstBook.Sheets("II.2.A").Range("RNG_II2A").Column - 5
        Set dstCell = DstBook.Sheets("Data").Range(dstCell, dstCell.Offset(curCnt))
        dstCell = ""
        Set dstCell = DstBook.Sheets("Data").Range("tblVillageStart")
        
        i = 1
        While Len(Trim(prcCell)) > 0
            dstCell = prcCell
            Set dstCell = dstCell.Offset(1)
            Set prcCell = prcCell.Offset(1)
            i = i + 1
        Wend
        
        ' now we have to modify village data
        ' curCnt is the number of existing village,
        ' 2 is the minimum number of village for each commune
        If i - 1 <= 2 Then i = 201 ' This is a very problem stuff
        ModifyColumns i - curCnt
        
        i = 100
        ' tell caller about the result
        ' Get data from II.2.A  dta_bsc_vil
        Application.StatusBar = MSG("MSG_COPY_II2A")
        CopyObject .Sheets("II.2.A"), SrcBook, "dta_bsc_vil", DstBook.Sheets("II.2.A"), "dta_bsc_vil", , False, True
        ' Now II.2 data
        ' Now II.2.B data
        ''dta_bsc_vil for II.2.A - then refer back to II.2
            'TBLMAJORINDS II.2.B Key indicators
            'II2BFIRST and II2BLAST for II.2.B
        
        SrcBook.Close False
    End With
    
    SuccessFullCall = True
StepEnd:
    On Error Resume Next
    If Not SrcBook Is Nothing Then SrcBook.Close False
    If Err.Number > 0 Then
        MsgBox MSG("MSG_UNKNOWN_ERROR"), vbCritical
        Err.Clear
    End If
    Set SrcBook = Nothing
    Set prcCell = Nothing
    Set DstBook = Nothing
StepExit:
    ' release all
    ShowOff True
End Sub

Private Sub CopyObject(SrcObj As Worksheet, SrcWrk As Workbook, SrcRangeName As String, _
        DstObj As Worksheet, DstRangeName As String, Optional IsTbl5A As Boolean = False, _
        Optional ShouldConvert As Boolean = True, Optional DirectCopy As Boolean = False)
    
    Dim SrcCurCell As Range, DstCurCell As Range, rngTarget As Range, rngSource As Range
    With SrcObj
        'hacked 21 june - release any filter
        'need to revise this::
        ShowAll SrcObj
        ' Sort the source sheet first
        If IsTbl5A Then SortTable SrcWrk, SrcObj.Name, SrcRangeName, "A6", "B6"
        
        ' Convert them all to Unicode
        If ShouldConvert Then ConvertRange .Range(SrcRangeName)
        If DirectCopy Then
            ' just copy across
            Set rngSource = .Range(SrcRangeName)
            Set DstCurCell = DstObj.Range(DstRangeName)
        Else
            ' Now just trying to find last row of the existing sheet
            Set DstCurCell = GetLastCell(DstObj.Range(DstRangeName).Cells(1, 1))
            
            ' now send it to the existing sheet
            Set rngTarget = .Range(SrcRangeName)
            Set SrcCurCell = rngTarget.Cells(1, 1)
            
            ' find last cell
            Set SrcCurCell = GetLastCell(SrcCurCell)
            ' Now just use a copy and paste
            Set rngSource = .Range(rngTarget.Cells(1, 1), SrcCurCell.Offset(, rngTarget.Columns.Count - 1))
            Set DstCurCell = DstCurCell.Offset(1)
        End If
        ' check number of column from both table and only copy what is there...
        ' Then try to revise
        If rngSource.Columns.Count < DstCurCell.Columns.Count Then
            'We are copying data from old version to this new
            ' Now we should have to copy column to column
            Dim i As Long, SrcCol As Range, DstCol As Range, DstColIndex As Long
            For i = rngSource.Columns.Count To 1 Step -1
                ' get back one row
                Set SrcCol = rngSource.Cells(1, i).Offset(-1)
                DstColIndex = FindColHeader(DstObj, SrcCol.Row, SrcCol.Value)
                If DstColIndex > 0 Then
                    Set DstCol = DstCurCell.Cells(SrcCol.Row, DstColIndex).Offset(-1)
                Else
                    ' this column has been discarded in the new version... what a mess - I have no done this before
                End If
            Next i
            GoTo ExitSub
        End If
        rngSource.Copy
        DstCurCell.PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End With
ExitSub:
    ' Release all
    Set SrcCurCell = Nothing
    Set DstCurCell = Nothing
    Set rngTarget = Nothing
    Set rngSource = Nothing
End Sub


Private Function GetLastCell(CellObj As Range) As Range
    While Len(Trim(CellObj)) > 0
        Set CellObj = CellObj.Offset(1)
    Wend
    Set GetLastCell = CellObj.Offset(-1)
End Function

Private Function FindColHeader(shtObj As Worksheet, FindRow As Long, FindTxt As String) As Long
    ' This function will return number of column with data specified in the Find text
    Dim FoundCell As Boolean, CellObj As Range, i As Long
    Set CellObj = shtObj.Cells(FindRow, 1)
    While i < 10 And Not FoundCell
        If Len(Trim(CellObj)) = 0 Then
            i = i + 1
        ElseIf CellObj = FindTxt Then
            FoundCell = True
        End If
        Set CellObj = CellObj.Offset(, 1)
    Wend
    If FoundCell Then FindColHeader = CellObj.Column - 1
End Function

Private Function ShrinkRange(rngIn As Range) As Range
    Dim LastCell As Range, tmpRange As Range
    Set LastCell = rngIn.Cells(rngIn.Rows.Count, 1)
    While Len(Trim(LastCell)) = 0
        Set LastCell = LastCell.Offset(-1)
    Wend
    Set tmpRange = rngIn.Range(rngIn.Cells(1, 1), LastCell)
    Set ShrinkRange = tmpRange
End Function

Private Sub ParseRange(frBook As Workbook, toBook As Workbook, shtName As String, RngName As String, Optional NeedUnprotect As Boolean = False)
    Dim RngArr As Variant, i As Long
    ' Revised by Ngoc on May 7 2014
    If NeedUnprotect Then XUnProtectSheet toBook.Sheets(shtName)
    RngArr = Split(RngName, ",")
    For i = 0 To UBound(RngArr)
        toBook.Sheets(shtName).Range(RngArr(i)) = frBook.Sheets(shtName).Range(RngArr(i))
    Next
    If NeedUnprotect Then XProtectSheet toBook.Sheets(shtName)
End Sub

Private Function RangeValid(RangeName As String, shtObj As Worksheet) As Boolean
    Dim txtRange As Range
    On Error GoTo errHandler
    Set txtRange = shtObj.Range(RangeName)
    RangeValid = True
errHandler:
End Function

Function SheetValid(SheetName As String, WrbObj As Workbook) As Boolean
    Dim txtRange As Worksheet
    On Error GoTo errHandler
    Set txtRange = WrbObj.Sheets(SheetName)
    SheetValid = True
errHandler:
End Function

Sub ModifyColumns(Optional NumberOfCols As Long = 1)
    'This is a hack to help people add/remove column for a new village
    ' First - unprotect the sheet
    XUnProtectSheet Sheet4
    
    Dim rngEnd As Range, rngStart As Range, i As Long
    Dim IsRemove As Boolean
    If NumberOfCols < 0 Then IsRemove = True
    ' Revert parametter for looping
    NumberOfCols = Abs(NumberOfCols)
    
    Set rngStart = Range("RNG_IIAST").Offset(0, 4)
    Set rngEnd = Range("RNG_II2A")
    
    Application.StatusBar = MSG("MSG_CREATE_II2A")
    
    ' After adding/removing - do the formatting
    For i = 1 To NumberOfCols
        If IsRemove Then
            ' Just not to allow deletion if there is only 02 columns left
            If rngEnd.Column - rngStart.Column <= 2 Then GoTo ExitCode
            Sheet4.Range("RNG_II2A").Offset(, -1).EntireColumn.Delete
        Else
            Sheet4.Range("RNG_II2A").EntireColumn.Insert
        End If
    Next
ExitCode:
    ' Now recreate fomular and stuff.
    CreateFomular
    FormatHeaderCell
    
    'Clean up
    Set rngStart = Nothing
    Set rngEnd = Nothing
    
    XProtectSheet Sheet4
End Sub

Private Sub CreateFomular()
    'This will help reformatting newly created table
    ' Begining column shall be total - 1 (coz RNG_IIA always stays at the end
    ' Range("RNG_IIAST").Offset(1, 4) offset 4 will always be the column for total
    ' Range("RNG_II2A_CELL_LAST").Offset(-1) will alway be the last cell at total column
    Dim rngStart As Range, rngEnd As Range, rngLastCell As Range
    Dim rngTotal As Range
    Dim MyCell As Range, i As Long
    
    ' reassign current worksheet
    Dim CurrentWorksheet As Worksheet
    Set CurrentWorksheet = ThisWorkbook.Sheets("II.2.A")
    
    With CurrentWorksheet
        Set rngStart = .Range("RNG_IIAST").Offset(0, 4)
        Set rngEnd = .Range("RNG_II2A")
        ' already inserted columns... so this failed
        Set rngLastCell = .Range("dta_bsc_vil").Cells(.Range("dta_bsc_vil").Rows.Count, 1).Offset(0, -1)
        Set rngTotal = .Range(rngStart.Offset(1), rngLastCell)
        
        ' Now that create total fomular
        ' + Create total column
        rngTotal.Formula = "=SUM(INDIRECT(""RC[1]" & ":RC[" & rngEnd.Column - rngStart.Column & "]"",FALSE))"
        
        ' + Create header link to data
        Set MyCell = Range("tblVillageStart")
        i = 0
        While Len(Trim(MyCell)) > 0
            i = i + 1
            .Range("RNG_IIAST").Offset(, 4 + i).Formula = "=INDIRECT(""Data!" & MyCell.Address & """)"
            Set MyCell = MyCell.Offset(1)
        Wend
        
        ' we have to unlock all data cells in this tables
        .Range("dta_bsc_vil").Locked = False
    End With
    Set rngStart = Nothing
    Set rngEnd = Nothing
    Set rngLastCell = Nothing
    Set rngTotal = Nothing
    Set MyCell = Nothing
    Set CurrentWorksheet = Nothing
End Sub

Private Sub FormatHeaderCell()
    ' Just for formatting the header
    Dim rngStart As Range, rngEnd As Range, MyCell As Range
    Set rngStart = Range("RNG_IIAST").Offset(0, 4)
    Set rngEnd = Range("RNG_II2A")
    Set MyCell = Sheet4.Range(rngStart.Offset(0, 1).Address & ":" & rngEnd.Address)
    
    ' Now format the header
    With MyCell
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 90
    End With
    With MyCell.Font
        .Name = "Times New Roman"
        .FontStyle = "Bold"
        .Size = 10
    End With
    With MyCell.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Set rngStart = Nothing
    Set rngEnd = Nothing
    Set MyCell = Nothing
End Sub

Function GetOpenWorkbook(FilePath As String) As Workbook
    'Open a workbook
    On Error GoTo errHandler
    Dim WrkBook As Workbook
    Set WrkBook = Application.Workbooks.Open(FilePath, False, True)
    Set GetOpenWorkbook = WrkBook
errHandler:
    Set WrkBook = Nothing
End Function

Sub RemoveProtection()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        XUnProtectSheet sh
    Next
    ThisWorkbook.Unprotect "d1nd1sk"
End Sub

Sub ListName()
    Dim sh As Worksheet, wrk As Workbook, theName As Name
    Set wrk = ThisWorkbook
    For Each theName In wrk.Names
        If theName.RefersToLocal Like "*II.2*" Then
            Debug.Print theName.Name
            'dta_bsc_vil for II.2.A - then refer back to II.2
            'TBLMAJORINDS II.2.B Key indicators
            'II2BFIRST and II2BLAST for II.2.B

        End If
    Next
    Set wrk = Nothing
End Sub

Function GetDistintiveList(TableName As String, KeyColumn As Long, ColumData As Long, Optional UseListOnly As Boolean = True) As String
    ' Check whether a temporary variable is valid
    On Error GoTo errHandler
    If IsCollection(CachedListDistinct) And OldTableName = TableName Then GoTo SetFuncValue
    OldTableName = TableName
    ' Now built the list, the sortable has been done before so we don't care
    Dim TheRange As Range, theCell As Range, StrCount As Long, SpStr As String
    
    Dim ColDistinctive As New Collection
    Set TheRange = ThisWorkbook.Names(TableName).RefersToRange
    SpStr = "[||]"
    
    Dim txtDistinct() As String, txtListing() As String, colCount As Long
    Dim i As Long, MaxPos As Long, MaxStr As String, xPos As Long
    
    With TheRange
        colCount = .Columns.Count
        ReDim txtDistinct(colCount - 1)
        ReDim txtListing(colCount - 1)
        ReDim ColListing(xPos)
        
        'We keep each stuff in one collection item and the very first shall alway be the type, frequency
        ' next hack - convert range to array for quicker access
        Set theCell = .Cells(1, 1)
        While theCell <> ""
            ' move through all column
            If ", " & SpStr & theCell.Offset(0, KeyColumn - 1) & SpStr <> txtListing(KeyColumn - 1) Then
                If i > 0 Then
                    ' if already in process, so flush current array to variable and startnew array
                    xPos = xPos + 1
                    ReDim Preserve ColListing(xPos)
                    For i = LBound(txtListing) To UBound(txtListing)
                        ColListing(xPos).Add Mid(Replace(txtListing(i), SpStr, ""), 3)
                    Next
                    ReDim txtListing(colCount - 1)
                End If
            End If
            For i = 1 To colCount
                If InStr(txtListing(i - 1), SpStr & theCell.Offset(0, i - 1) & SpStr) = 0 Then
                    ' only add the new thing
                    txtListing(i - 1) = txtListing(i - 1) & ", " & SpStr & theCell.Offset(0, i - 1) & SpStr
                End If
                If InStr(txtDistinct(i - 1), SpStr & theCell.Offset(0, i - 1) & SpStr) = 0 Then
                    txtDistinct(i - 1) = txtDistinct(i - 1) & ", " & SpStr & theCell.Offset(0, i - 1) & SpStr
                    If i = KeyColumn Then
                        StrCount = 1
                        If MaxPos < StrCount Then
                            MaxPos = StrCount
                            MaxStr = theCell.Offset(0, KeyColumn - 1)
                        End If
                    End If
                Else
                    ' Find Max freq for key column
                    If i = KeyColumn Then
                        StrCount = StrCount + 1
                        If MaxPos < StrCount Then
                            MaxPos = StrCount
                            MaxStr = theCell.Offset(0, KeyColumn - 1)
                        End If
                    End If
                End If
            Next
            ' Okie - get along all columns already, now we need to see whether the next would be different
            Set theCell = theCell.Offset(1)
        Wend
        ' Add the last stuff
        xPos = xPos + 1
        ReDim Preserve ColListing(xPos)
        
        ' Now pass the array to the collection and cached them
        For i = LBound(txtDistinct) To UBound(txtDistinct)
            ColDistinctive.Add Mid(Replace(txtDistinct(i), SpStr, ""), 3)
            ColListing(xPos).Add Mid(Replace(txtListing(i), SpStr, ""), 3)
        Next
        Set CachedListDistinct = ColDistinctive
    End With
    ' Now we have to find the most appeared object
    For i = 1 To UBound(ColListing)
        If ColListing(i).Item(KeyColumn) = MaxStr Then
            Set ColListing(0) = ColListing(i)
            If i < UBound(ColListing) Then Set ColListing(i) = ColListing(i + 1)
        End If
    Next
    ' resize the array
    ReDim Preserve ColListing(i - 2)
SetFuncValue:
    If UseListOnly Then
        GetDistintiveList = ColListing(CurrentPointer).Item(ColumData)
    Else
        GetDistintiveList = CachedListDistinct(ColumData)
    End If
errHandler:
    If Err.Number <> 0 Then Debug.Print Err.description & "CurrentPointer=[" & CurrentPointer & "]"
End Function

Sub TestAccessFormD()
    Set CachedListDistinct = Nothing
    ReDim ColListing(0)
    SortTable ThisWorkbook, "II.5.D", "tblUnicode_4", "C6", "A6"
    'CurrentPointer = 0
    'Debug.Print GetDistintiveList("tblUnicode_4", 2, 2) & "//" & GetDistintiveList("tblUnicode_4", 2, 1) & "//" & GetDistintiveList("tblUnicode_4", 2, 6)
    For CurrentPointer = 0 To UBound(ColListing)
        Debug.Print GetDistintiveList("tblUnicode_4", 2, 2)
    Next
End Sub

Function GetOption(TxtIn As String) As ObjectEquation()
    ' This will read the parametter and convert into an array for later processing
    Dim MyObj() As ObjectEquation, i As Long, ArrItem As Variant
    Dim myArr As Variant
    myArr = Split(TxtIn, "/")
    ReDim MyObj(UBound(myArr))
    For i = LBound(myArr) To UBound(myArr)
        ArrItem = Split(myArr(i), "=")
        With MyObj(i)
            .VariableName = ArrItem(0)
            .VariableFomular = Evaluate(ArrItem(1))
        End With
    Next
    GetOption = MyObj
End Function

Function CountMaxRepetition(RangeName As String, CountColumn As Long, _
    Optional ReferColumn As Long = 0, Optional CountOnly As Long = 1, Optional InsertLineBreak As Boolean = False) As Variant
    'This function will count and get the maximum number of object repetition
    Dim TheRange As Range, theCell As Range, RetObj As TextObject
    Dim StrTxt As String, StrCount As Long, MaxPos As Long, MaxStr As String, MaxRefText As String
    Dim RefStrTxt As String
 
    Set TheRange = ThisWorkbook.Names(RangeName).RefersToRange
    ' Turn the range to an array for quick access
    Set theCell = TheRange.Cells(1, CountColumn)
    While theCell <> ""
        If StrTxt <> theCell Then
            StrTxt = theCell
            If ReferColumn <> 0 Then RefStrTxt = theCell.Offset(, ReferColumn - CountColumn)
            StrCount = 1
        Else
            StrCount = StrCount + 1
            If ReferColumn <> 0 Then RefStrTxt = RefStrTxt & "[SEP]" & theCell.Offset(, ReferColumn - CountColumn)
            If MaxPos < StrCount Then
                MaxPos = StrCount
                MaxStr = StrTxt
                MaxRefText = RefStrTxt
            End If
        End If
        Set theCell = theCell.Offset(1)
    Wend
    
    'On Error Resume Next
    Select Case CountOnly
    Case 1:
        CountMaxRepetition = MaxStr
    Case 2:
        CountMaxRepetition = IIf(InsertLineBreak, Replace(MaxRefText, "[SEP]", vbCrLf), Replace(MaxRefText, "[SEP]", ", "))
    Case 3:
        CountMaxRepetition = Replace(MaxRefText, "[SEP]", ",")
    End Select
End Function

Private Function GetAverage(inputText As String) As String
    On Error GoTo errHandler
    Dim i As Long, theText As String, myArr As Variant, theTotal As Double
    theText = Replace(Replace(Replace(inputText, "(", ""), ")", ""), " ", "")
    myArr = Split(theText, ",")
    For i = LBound(myArr) To UBound(myArr)
        theTotal = theTotal + CDbl(myArr(i))
    Next
    theTotal = theTotal / i
    GetAverage = theTotal
errHandler:
End Function

Function IsCollection(inCol As Object) As Boolean
    ' Check whether an object is a collection or not
    On Error GoTo errHandler
    IsCollection = IIf(inCol.Count > 0, True, False)
errHandler:
End Function
