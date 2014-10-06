Option Explicit

Function MSG(MsgName As String) As String
    ' This function will return expected string for better userinterface
    MSG = "False"
    Dim MyCell As Range, FoundObj As Boolean
    Set MyCell = ThisWorkbook.Sheets("Data").Range("MSG_ID_START").Offset(1)
    While Not FoundObj
        If Len(Trim(MyCell)) <= 0 Then
            FoundObj = True
        Else
            If MyCell = MsgName Then
                FoundObj = True
                MSG = MyCell.Offset(, 1)
            End If
        End If
        Set MyCell = MyCell.Offset(1)
    Wend
End Function

Function FalseInput(CtrlName As Control) As Boolean
    Dim tData As String
    If CtrlName = "" Then Exit Function
    If Not IsDate(CtrlName) Then GoTo tCont
    tData = InputDate(CtrlName)
    If Not tData Like "12:00*" Then Exit Function
tCont:
    CtrlName = ""
    CtrlName.SetFocus
    FalseInput = True
End Function

Function InputDate(iDateStr As Variant) As Date
    ' Send data piece from database to console
    ' default the data will from db to console, output shall be formated
    ' input shall be converted back to serial date
    Dim iStr As String, iSpliter As Variant
    
    On Error GoTo errHandler
    iSpliter = Split(iDateStr, "/")
    If UBound(iSpliter) < 2 Then GoTo errHandler
    ' Now we have to see what locale we are now at
    InputDate = DateSerial(iSpliter(2), iSpliter(0), iSpliter(1))
errHandler:
End Function

Function GetFolder(strPath As String, Optional FilePicker As Boolean = False, Optional FileExtension As String = "*.*") As String
    Dim fldr As FileDialog
    Dim sItem As String
    If FilePicker Then
        Set fldr = Application.FileDialog(msoFileDialogFilePicker)
    Else
        Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    End If
StepResumeFolder:
    With fldr
        .Title = MSG("MSG_SELECTDATAFOLDER")
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If FilePicker Then
            .Filters.Clear
            .Filters.Add "Mirosoft Excel File", FileExtension
        End If
        If .Show <> -1 Then
            'user select cancel
            sItem = ""
        Else
            sItem = .SelectedItems(1)
        End If
    End With
    
    ' Test to make sure that user selected anything
    Dim FileLocation As String, FldBrowser As String
    
    If Not FileOrDirExists(sItem, FilePicker) Then
        If MsgBox(MSG("MSG_SELECT_NO_FILE"), vbInformation + vbOKCancel) = vbOK Then GoTo StepResumeFolder
        ' safe exit
        sItem = ""
        GoTo NextCode
    End If
    If FilePicker Then
        ' User select current file or not?
        If sItem = ThisWorkbook.Path & "\" & ThisWorkbook.Name Then
            MsgBox MSG("MSG_ERROR_THIS_FILE"), vbInformation
            GoTo StepResumeFolder
        End If
    End If
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Function FileOrDirExists(PathName As String, Optional FileObject As Boolean = False) As Boolean
'No need to set a reference if you use Late binding
    Dim FSO As Object
    Dim FilePath As String, lRet As Boolean

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If PathName = "" Then Exit Function
    If FileObject Then
        FileOrDirExists = FSO.FileExists(PathName)
    Else
        FileOrDirExists = FSO.FolderExists(PathName)
    End If
    Set FSO = Nothing
End Function

Property Get AppDecimal() As String
    ' return application locale
    If InStr(Format("12345", "#,##0"), ",") > 0 Then AppLocale.DecimalSeparator = "."
End Property

Property Get ListSeparator() As String
    ' return application locale
    If Application.version > 14 Then ListSeparator = ";" Else ListSeparator = ","
End Property
