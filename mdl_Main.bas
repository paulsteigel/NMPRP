Option Explicit

' For storing item attribute
Private Type ItemAttributes
    ItemDetails As String
    ItemHeading As String
    ItemEmphasizeHeading As String
    DataSource As String
    Name As String
End Type

Private Type TagItem
    TagOpen As String
    TagClose As String
End Type
Private myWordApp As Object ' for word application

Sub GenerateDocs()
    'RegisterAction
    ShowStatus ""
    ' First - convert all to Unicode
    ShowOff False
    
    Dim myWordApp As Object, myWordDoc As Object, LocalSetting As String, RplStr As String, DocStr As String
    
    LocalSetting = AppDecimal
    
    
    Set myWordDoc = CreateWordDocument(myWordApp)
    myWordApp.Visible = False
    
    Dim i As Long, HasWordError As Boolean
    Dim MyRange As Range
    Set MyRange = Range("PROPOSAL")
    RplStr = Range("TAB_OBJ")
    
    ' now generate all style
    ShowStatus MSG("MSG_CREATE_STYLES")
    HasWordError = GenerateWordStyle(myWordDoc, myWordApp)

    If HasWordError Then GoTo errHandler
    Dim j As Long, FieldItems As Variant
    
    Dim AllRowCount As Long
        
    ' Style for output
    Dim DocStyle As ItemAttributes, tmpValue As String
    Dim theSheet As Worksheet, tmpRange As Range
    
    Dim MsgPasstoWord As String, MsgFormatTable As String, MsgProccesingWordDocs As String, MsgFinished As String
    MsgPasstoWord = MSG("MSG_PASS_SECTOR_TO_WORD")
    MsgFormatTable = MSG("MSG_PROCESS_TABLE")
    MsgProccesingWordDocs = MSG("MSG_PROCESS_TABLE_DATA")
    MsgFinished = MSG("MSG_FINISHED")
       
    With myWordDoc
        AllRowCount = MyRange.Rows.Count
        For i = 1 To AllRowCount
            ' First get style
            If MyRange.Cells(i, 4) <> "" Then
                DocStyle.ItemHeading = MyRange.Cells(i, 4)
            Else
                DocStyle.ItemHeading = "Normal"
            End If
                                
            ' Just move from the begining to the end and apply thing...
            For j = 1 To 3
                If Not IsError(MyRange.Cells(i, j)) Then
                    If MyRange.Cells(i, j) <> "" Then
                        
                        DocStr = CStr(MyRange.Cells(i, j))
                            
                        ' Now add data
                        If MyRange.Cells(i, j) Like "FIELD::*" Then
                            FieldItems = Split(Replace(DocStr, "FIELD::", ""), "/")
                            'FIELD::TITLE[ANN_T_03]/TABLE[ANNEX_03]/FILTER[1]
                            
                            'Table title
                            FieldItems(0) = Replace(Replace(FieldItems(0), "TITLE[", ""), "]", "")
                            'Table Range
                            FieldItems(1) = Replace(Replace(FieldItems(1), "TABLE[", ""), "]", "")
                            'Filter column
                            FieldItems(2) = Replace(Replace(FieldItems(2), "FILTER[", ""), "]", "")
                            'Insert table direct into main text
                            
                            DocStr = Range(FieldItems(0))
                            DocStyle.ItemHeading = "Caption"
                            
                            ' Insert table caption
                            InsertPara myWordDoc, DocStyle, DocStr
                            
                            ' Insert table
                            Set tmpRange = Range(FieldItems(1))
                            Set theSheet = tmpRange.Parent
                            XUnProtectSheet theSheet
                            
                            ' set filter first
                            theSheet.Range(FieldItems(1)).AutoFilter FIELD:=Val(FieldItems(2)), Criteria1:="<>"
                            
                            tmpRange.Copy
                            
                            '.Paragraphs.Add
                            '.Paragraphs.Last.Style = "NoFirstLine"
                            .Paragraphs.Last.Range.Paste
                            
                            Application.CutCopyMode = False
                            
                            'Release filter
                            theSheet.ShowAllData
            
                            Set tmpRange = Nothing
                            XProtectSheet theSheet
                            Set theSheet = Nothing
                        Else
                            ' Just normal text
                            If DocStyle.ItemHeading Like "Heading*" Then
                                ' Remove numbering stuff
                                DocStr = Mid(DocStr, InStr(DocStr, " ") + 1)
                            Else
                                DocStr = Replace(DocStr, RplStr, "")
                            End If
                            ' Turn up the uppercase
                            If DocStyle.ItemHeading = "Title" Then DocStr = UCase(DocStr)
                            
                            InsertPara myWordDoc, DocStyle, DocStr
                        End If
                        Exit For
                    End If
                End If
            Next
            ShowStatus MsgPasstoWord & " " & Format((i - 2) * 100 / AllRowCount, "##0") & "% " & MsgFinished
        Next
        
        ' formatt some specific texts
        RemoveTagAndFormat myWordDoc
        
        ' Step 2: Insert Annexes
        Set MyRange = Range("LST_ANNEX")
        Set MyRange = MyRange.Cells(1)
        
        While MyRange <> ""
            Set theSheet = GetSheet(MyRange.Offset(0, 1))
            
            XUnProtectSheet theSheet
            
            ' set filter first
            If MyRange.Offset(0, 5) <> "" Then
                theSheet.Range(MyRange.Offset(0, 5)).AutoFilter FIELD:=1, Criteria1:="<>"
            End If
            Range(CStr(MyRange)).Copy
            ' In
            InsertSection myWordDoc
            
            If MyRange.Offset(0, 3) <> "" Then
                .Paragraphs.Last.Range.Text = Range(CStr(MyRange.Offset(0, 3)))
                .Paragraphs.Last.Style = "Phuluc"
            End If
            If MyRange.Offset(0, 4) <> "" Then
                'a sub table needed
                .Paragraphs.Add
                .Paragraphs.Last.Style = "Phuluc_sub"
                .Paragraphs.Last.Range.Text = Range(CStr(MyRange.Offset(0, 4)))
            End If
            .Paragraphs.Add
            .Paragraphs.Last.Style = "NoFirstLine"
            
            '.paragraphs.Last.Range.PasteSpecial DataType:=wdPasteRTF
            .Paragraphs.Last.Range.Paste
            
            ' Set current section layout
            SetSectionLayout myWordDoc, IIf(MyRange.Offset(0, 2) = 1, True, False)
            
            Application.CutCopyMode = False
            'Release filter
            If MyRange.Offset(0, 5) <> "" Then theSheet.ShowAllData
            Set tmpRange = Nothing
            
            XProtectSheet theSheet
            
            Set MyRange = MyRange.Offset(1)
            ShowStatus MsgProccesingWordDocs & " " & MyRange
        Wend
        
        Set theSheet = Nothing
        ReformatWordTable myWordDoc, MsgProccesingWordDocs, MsgFormatTable, MsgFinished
    End With
        
errHandler:
    If HasWordError Then
        Err.Clear
        MsgBox MSG("MSG_WORD_NOT_CLOSE"), vbCritical
    Else
        ShowStatus MSG("MSG_CREATE_BUSINESS_PLAN")
    End If
    ShowOff True
    myWordApp.Visible = True
    myWordApp.Activate
    Set myWordDoc = Nothing
    Set myWordApp = Nothing
End Sub

Private Function GetSheet(WildCardName As String) As Worksheet
    Dim Sh As Worksheet
    For Each Sh In ThisWorkbook.Sheets
        If InStr(Sh.Name, WildCardName) <> 0 Then
            Set GetSheet = Sh
            Exit For
        End If
    Next
End Function

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
        prCount = .Paragraphs.Count
        .Paragraphs(prCount).Range.Style = .Styles(ItemStyle.ItemHeading)
        .Paragraphs(prCount).Range.Text = ItemText
        
        ' Add a new prg...
        .Paragraphs.Add
        If ItemStyle.ItemDetails <> "" And Not OverideAdd Then
            ' Add new introduction line if neccessary
            tmpItem.ItemHeading = tmpItem.ItemEmphasizeHeading
            tmpText = tmpItem.ItemDetails
            tmpItem.ItemDetails = ""
            InsertPara DocObj, tmpItem, tmpText
        End If
    End With
End Sub

Private Function CountTable(Obj As Object) As Long
    On Error GoTo errHandler
    CountTable = Obj.Tables.Count
errHandler:
End Function

Sub XProtectSheet(s As Worksheet)
    s.Protect "d1ndh1sk", Contents:=True, AllowFormattingCells:=False, AllowFormattingColumns:=True, DrawingObjects:=True, Scenarios:=True, _
    AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, UserInterfaceOnly:=True
End Sub

Sub XUnProtectSheet(s As Worksheet)
    s.Unprotect "d1ndh1sk"
End Sub

Private Function GetLastCell(CellObj As Range) As Range
    While Len(Trim(CellObj)) > 0
        Set CellObj = CellObj.Offset(1)
    Wend
    Set GetLastCell = CellObj.Offset(-1)
End Function

Private Function FindColHeader(shtObj As Worksheet, FindRow As Long, FindTxt As String) As Long
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

Private Function SheetValid(SheetName As String, WrbObj As Workbook) As Boolean
    Dim txtRange As Worksheet
    On Error GoTo errHandler
    Set txtRange = WrbObj.Sheets(SheetName)
    SheetValid = True
errHandler:
End Function

Function GetOpenWorkbook(FilePath As String) As Workbook
    'Open a workbook and disable macro...
    On Error GoTo errHandler
    Dim WrkBook As Workbook
    'Dim secAutomation As MsoAutomationSecurity
    'secAutomation = Application.AutomationSecurity
    'Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.EnableEvents = False
    Set WrkBook = Application.Workbooks.Open(FilePath, False, True)
    'Application.AutomationSecurity = secAutomation
    Application.EnableEvents = True
    Set GetOpenWorkbook = WrkBook
errHandler:
    Set WrkBook = Nothing
End Function

Sub ProtectObject(Optional ProtectEnabled As Boolean = False)
    Dim Sh As Worksheet
    If ProtectEnabled Then ThisWorkbook.Protect "d1nd1sk" Else ThisWorkbook.Unprotect "d1nd1sk"
    For Each Sh In ThisWorkbook.Sheets
        If ProtectEnabled Then XProtectSheet Sh Else XUnProtectSheet Sh
    Next
End Sub

Sub ClearData()
    Dim Sht As Worksheet, theCell As Range
    Dim ExName As String, OldSetting As String
    If MsgBox(MSG("MSG_DELETE"), vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    ' Exclusion range
    ExName = "Main,Data,Dexuat,PL2"
    ShowOff
    DoEvents
    OldSetting = Range("COND_GOV_OPT")
    For Each Sht In ThisWorkbook.Sheets
        If InStr(ExName, Sht.Name) = 0 Then
            For Each theCell In Sht.Range("Print_Area").Cells
                If Not theCell.Locked Then
                    'theCell.NumberFormat = "General"
                    If Not theCell.FormulaHidden Then theCell = Null
                End If
            Next
        End If
    Next
    Range("COND_GOV_OPT") = OldSetting
    ShowOff True
    ShowStatus MSG("MSG_FINISHED")
    ' reset some objects
    CreateSampleX True
End Sub

Function HasName(InCell As Range, CheckName As String) As Boolean
    On Error GoTo errHandler
    If InCell.Name = CheckName Then HasName = True
errHandler:
End Function

Sub CreateSample()
    CreateSampleX
End Sub

Private Sub CreateSampleX(Optional CleanData As Boolean = False)
    If CleanData Then GoTo ResumeStep
    If MsgBox(MSG("CREATE_SAMPLE"), vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
ResumeStep:
    Dim Sht As Worksheet, theCell As Range, theName As Name
    Dim ExName As String, CellStart As Range
    ' Exclusion range
    
    Set CellStart = Range("NAME_SAMPLE").Offset(1)
    ShowOff
    DoEvents
    ' don't do anything like 9
    While CellStart <> ""
        If CellStart.Offset(0, 1) <> 9 Then
            If Not Range(CellStart).Locked Then
                If CleanData Then
                    Range(CellStart) = Null
                Else
                    Range(CellStart) = Range(CellStart.Offset(0, 2), CellStart.Offset(0, Range(CellStart).Columns.Count + 1)).Value
                End If
            End If
        End If
        Set CellStart = CellStart.Offset(1)
    Wend
    ShowOff True
End Sub

Sub RetrieveSampleData()
    ' This will help collecting a new set of data for sampling...
    Dim Sht As Worksheet, theCell As Range, theName As Name
    Dim ExName As String, CellStart As Range
    ' Exclusion range
    ExName = "Main,Data,Dexuat" ',PL2,T12-PL5,T10-11-PL4,T9-PL3,T3-PL1"
    Set CellStart = Range("NAME_SAMPLE").Offset(1)
    ShowOff
    DoEvents
    For Each theName In ThisWorkbook.Names
        'If InStr(ExName, Sht.Name) = 0 Then
        '    For Each theName In Sht.Names 'Range("Print_Area").Range
            'On Error Resume Next
            'If InStr(theName, "#REF") <> 0 Then
            'Else
                If Not Range(theName).Locked Then
                    CellStart = theName.Name  '.Name '.Address(External:=True)
                    CellStart.Offset(0, 1) = Range(theName).Address(External:=True)
                    If Range(theName).Rows.Count = 1 Then
                        Range(theName).Copy
                        CellStart.Offset(0, 2).PasteSpecial xlPasteValues
                    End If
                    Set CellStart = CellStart.Offset(1)
                End If
            'End If
            
        '    Next
        'End If
    Next
    ShowOff True
End Sub

Sub EditCaption()
    Dim Sh As Shape, curSheet As Worksheet
    Set curSheet = Sheets("Dexuat")
    Set Sh = curSheet.Shapes("Button 77")
    If Sh.TextFrame.Characters.Text = MSG("MSG_CAP_SAVE") Then
        XUnProtectSheet curSheet
        ' Just block it now and save
        Sh.TextFrame.Characters.Text = MSG("MSG_EDIT_CAP")
        
        curSheet.Range("PRP_CAPTION").Locked = True
        XProtectSheet ActiveSheet
        GoTo ExitMe
    End If
    If MsgBox(MSG("EDIT_SENTENCE_CAP"), vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    XUnProtectSheet curSheet
    curSheet.Range("PRP_CAPTION").Locked = False
    Sh.TextFrame.Characters.Text = MSG("MSG_CAP_SAVE")
    
ExitMe:
    Set Sh = Nothing
    XProtectSheet curSheet
End Sub

Sub MergeData()

    ' This procedure shall help merging data from various table into this.
    ' By doing this, the application shall ask user from verifying some key question to make sure that they will not
    ' try to duplicate the import
    MsgBox MSG("MSG_IMPORT_LIMITED"), vbInformation
    '-----------------------------------------------------------------------
    
    ShowOff
    ' First - convert all to Unicode
    Dim SrcBook As Workbook
    Dim DstBook As Workbook
    Dim CellStart As Range, theCellSrc As Range, theCellDst As Range, i As Long, LstArray As Variant
    Dim ObjDlg As Dialog
    ' Now open the existing workbook to import data
    Set ObjDlg = Application.Dialogs.Item(xlDialogOpen)
    
    Dim FileLocation As String, FldBrowser As String

    FldBrowser = GetFolder(ThisWorkbook.Path, True, "*.xls")
    If FldBrowser = "" Then GoTo StepEnd
        
    On Error GoTo StepEnd
    If FldBrowser = "" Then GoTo StepExit
    Set SrcBook = GetOpenWorkbook(FldBrowser)
    
    Set DstBook = ThisWorkbook
    ' check if this file is good or not
    If Not SheetValid("Data", SrcBook) Or Not SheetValid("T1", SrcBook) Or Not SheetValid("T2", SrcBook) Then
        MsgBox MSG("MSG_BAD_FILE"), vbCritical
        GoTo StepEnd
    End If
        
    With SrcBook
        Application.StatusBar = MSG("MSG_COPY_DATA")
        
        ' Just move around data with old name...
        Set CellStart = DstBook.Names("NAME_SAMPLE").RefersToRange.Offset(1)
        
        DoEvents
        While CellStart <> ""
            ' For sing row
            Select Case Val(CellStart.Offset(0, 1))
            Case 1, 3:
                ' Loop until end, since there were a mistake in making name for the
                ' direct investment, we must change this a bit... by offset next 5 columns
                Set theCellSrc = SrcBook.Names(CStr(CellStart)).RefersToRange
                Set theCellDst = DstBook.Names(CStr(CellStart)).RefersToRange
            
                While Not theCellSrc.Locked
                    ' move all along cells...
                    For i = 1 To theCellSrc.Cells.Count
                        theCellDst.Cells(i).Value = theCellSrc.Cells(i).Value
                    Next
                    ' Next offset 5 columns
                    If Val(CellStart.Offset(0, 1)) = 3 Then
                        For i = 1 To 2
                            theCellDst.Offset(0, 5).Cells(i).Value = theCellSrc.Offset(0, 5).Cells(i).Value
                        Next
                    End If
                    Set theCellSrc = theCellSrc.Offset(1)
                    Set theCellDst = theCellDst.Offset(1)
                Wend
            Case 9:
                ' Loop until end, since there were a mistake in making name for the
                ' direct investment, we must change this a bit... by offset next 5 columns
                Set theCellSrc = SrcBook.Names(CStr(CellStart)).RefersToRange
                Set theCellDst = DstBook.Names(CStr(CellStart)).RefersToRange
                
                ' Copy a range
                For i = 1 To theCellSrc.Cells.Count
                    If Not theCellSrc.Cells(i).Locked Then
                        theCellDst.Cells(i).Value = theCellSrc.Cells(i).Value
                    End If
                Next
                
            Case 0:
                ' Just copy single cell
                If Not SrcBook.Names(CStr(CellStart)).RefersToRange.Locked Then
                    DstBook.Names(CStr(CellStart)).RefersToRange.Value = SrcBook.Names(CStr(CellStart)).RefersToRange.Value
                End If
            End Select
            
            Set CellStart = CellStart.Offset(1)
        Wend
        Application.StatusBar = MSG("MSG_COPY_DATA_LIST")
        ' now coppy all stuff
        LstArray = Split("LST_UNITS,LST_TRAIN_TYPE,PRO_UNIT,LST_TRAIN_TYPE,LST_PROCU_TYPE,LST_OPTION,LST_LOST_TYPE", ",")
        For i = LBound(LstArray) To UBound(LstArray)
            Set theCellSrc = SrcBook.Sheets("Data").Range(CStr(LstArray(i))).Cells(1)
            Set theCellDst = DstBook.Sheets("Data").Range(CStr(LstArray(i))).Cells(1)
                
            While theCellSrc <> ""
                If Not theCellSrc.Locked Then theCellDst.Value = theCellSrc.Value
                Set theCellDst = theCellDst.Offset(1)
                Set theCellSrc = theCellSrc.Offset(1)
            Wend
        Next
    End With
    SrcBook.Close False
    MsgBox Replace(MSG("MSG_FINISHED_MERGING"), "%REL%", "[" & FldBrowser & "]"), vbInformation
    
StepEnd:
    On Error Resume Next
    If Not SrcBook Is Nothing Then SrcBook.Close False
    If Err.Number > 0 Then
        MsgBox MSG("MSG_UNKNOWN_ERROR"), vbCritical
        Err.Clear
    End If
    Set SrcBook = Nothing
    Set DstBook = Nothing
    Set theCellDst = Nothing
    Set theCellSrc = Nothing
    Set CellStart = Nothing
StepExit:
    ' release all
    ShowOff True
End Sub

Private Sub createDbs()
    Dim CellStart As Range
    Dim dbRange As Range
    Set dbRange = Range("dbs")
    Set CellStart = Range("NAME_SAMPLE").Offset(1)
    ShowOff
    DoEvents
    ' don't do anything like 9
    While CellStart <> ""
        If CellStart.Offset(0, 1) <> 9 Then
            If Not Range(CellStart).Locked Then
                dbRange = CellStart.Value
            End If
        End If
        Set CellStart = CellStart.Offset(1)
        Set dbRange = dbRange.Offset(1)
    Wend
    ShowOff True
End Sub


'=====================================================================
'SOME NEW THINGS FOR NOTHING....CREATED ON JUNE 4 2014
'=====================================================================
' New updates
' Helping the ability of saving several proposals in one file
' Step 1: from the range, convert to an array type string and store in a column named with proposal
Private Function Array2Range(InputData As String) As Boolean
    ' First, create an array from inputData
    
End Function

Sub test2DB()
    Forms2Db
End Sub

Private Sub Forms2Db(Optional RecordName As String = "")
    ' this will help parsing form data to db
    If RecordName = "" Then
        MsgBox Replace(MSG("MSG_SAVE_2_DB"), "%REL%", Range("T_10_1")), vbInformation
        GetFormsData 1
    Else
        ' This is the edit mode, try to look for the current active profile...
        'ACT_EDT_COL
    End If
    ' Save data using T_10_1,T_1_PRV,T_1_DST,T_1_CMN,T_1_VIL
End Sub

Private Sub GetFormsData(ColId As Long)
    Dim Sht As Worksheet, theCell As Range, theName As Name
    Dim ExName As String, CellStart As Range
    Dim varObj As Variant
    
    ' Exclusion range
    Set CellStart = Range("dbs").Offset(1)
    ShowOff
    DoEvents
    ' don't do anything like 9
    While CellStart <> ""
        If Not Range(CellStart).Locked Then
            If Range(CellStart).Columns.Count > 1 Then
                ' convert source range to array
                varObj = Application.Transpose(Application.Transpose(Range(CellStart).Value))

                ' and parse array data to db
                CellStart.Offset(0, ColId) = Join(varObj, "[]")
            Else
                CellStart.Offset(0, ColId) = Range(CellStart).Value
            End If
        End If
        Set CellStart = CellStart.Offset(1)
    Wend
    Set CellStart = Nothing
    ShowOff True
End Sub
