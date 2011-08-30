Option Explicit
Option Base 0

'DisabledUpdatesErrors
'Turns Off Screen Updating and events and alerts for 
'better performance
'
'Usage
'Activate at the beginning of a sub and deactivate at the end
'DisableUpdatesErrors(True, "Please Wait, Working...")
''... do something
'DisableUpdatesErrors(False)

Sub DisableUpdatesErrors(Disabled As Boolean, Optional StatusBarMessage as String = "Working..." )
    On Error Resume Next
    With Excel.Application
        If Disabled = True Then
            .Cursor = xlWait
            .StatusBar = StatusBarMessage
            .EnableCancelKey = xlInterrupt
            .ScreenUpdating = False
            .EnableEvents = False
            .DisplayAlerts = False
        Else
			.Cursor = xlDefault
			.StatusBar = False
			.EnableCancelKey = xlInterrupt
			.ScreenUpdating = True
			.EnableEvents = True
        End If
    End With
End Sub



'LastRow
'Finds a the last row, either in the sheet given
'or in the current sheet if nothing is specified
'Gives predictable results - 0 if the sheet is empty
'The number of the last row containing data
'
'Usage
'Just call it when the last row is needed
'For i = 1 to LastRow
'	'... do somenthing
'Next

Function LastRow(Optional SheetName As String)
	'Disable Errors - If the sheet is completely empty will
	'be an error which will be put back to 0
    On Error Resume Next
    If SheetName = "" Then SheetName = ActiveSheet.Name
    LastRow = Sheets(SheetName).Cells.Find("*", SearchOrder:=xlByRows, LookIn:=xlValues, SearchDirection:=xlPrevious).EntireRow.Row
    If LastRow = "" Then LastRow = 0
End Function



'LastColumn
'Finds a the last column, either in the sheet given
'or in the current sheet if nothing is specified
'Gives predictable results - 0 if the sheet is empty
'The number of the last column containing data
'
'Usage
'Just call it when the last column is needed
'For i = 1 to LastColumn
'	'... do somenthing
'Next

Function LastColumn(Optional SheetName As String)
	'Disable Errors - If the sheet is completely empty will
	'be an error which will be put back to 0
    On Error Resume Next
    If SheetName = "" Then SheetName = ActiveSheet.Name
    LastColumn = Sheets(SheetName).Cells.Find("*", SearchOrder:=xlByColumns, LookIn:=xlValues, SearchDirection:=xlPrevious).EntireColumn.Column
    If LastColumn = "" Then LastColumn = 0
End Function



'DeleteEmptyRows
'Will delete all the empty rows in the selected range. If no range is selected will
'delete all the empty rows from the whole sheet
'
'Usage
'Call it in a sub
'DeleteEmptyRows
'
'Or bind it either to a button or a shortcut

Sub DeleteEmptyRows()
    Dim Rng             As Excel.Range
    Dim i               As Long
    On Error GoTo Err_Handl
    With ActiveSheet
        If Selection.Cells.Count = 1 Then
            Set Rng = Range(Rows(1), _
                Rows(.Cells.Find("*", SearchOrder:=xlByRows, LookIn:=xlValues, SearchDirection:=xlPrevious).Row))
        ElseIf Selection.Rows.Count = 1 Then
            If Application.CountA(Selection.EntireRow) = 0 Then Selection.EntireRow.Delete
                Exit Sub
        Else
            Set Rng = Selection
        End If
        For i = Rng.Rows.Count To 1 Step -1
            If Application.CountA(Rng.Rows(i).EntireRow) = 0 Then
                Rng.Rows(i).EntireRow.Delete
            End If
        Next i
    End With
    Exit Sub
Err_Handl:
        MsgBox "Error - Sheet might contain only empty rows", vbOKOnly
End Sub


'DeleteEmptyColumns
'Will delete all the empty columns in the selected range. If no range is selected will
'delete all the empty columns from the whole sheet
'
'Usage
'Call it in a sub
'DeleteEmptyColumns
'
'Or bind it either to a button or a shortcut

Sub DeleteEmptyColumns()
    Dim Rng             As Excel.Range
    Dim i               As Long
    On Error GoTo Err_Handl
    With ActiveSheet
        If Selection.Cells.Count = 1 Then
            Set Rng = Range(Columns(1), Columns(.Cells.Find("*", SearchOrder:=xlByColumns, LookIn:=xlValues, SearchDirection:=xlPrevious).Column))
        ElseIf Selection.Columns.Count = 1 Then
            If Application.CountA(Selection.EntireColumn) = 0 Then Selection.EntireColumn.Delete
            Exit Sub
        Else
            Set Rng = Selection
        End If
        For i = Rng.Columns.Count To 1 Step -1
            If Application.CountA(Rng.Columns(i).EntireColumn) = 0 Then
                Rng.Columns(i).EntireColumn.Delete
            End If
        Next i
    End With
    Exit Sub
Err_Handl:
        MsgBox "Error - Sheet might contain only empty columns", vbOKOnly
End Sub



'DeleteRowsCriteria
'Delete Rows with a certain value in them
'
'

Sub DelRowCrit(sValue As String, Optional Rng As Range)
    Dim i As Integer
    Dim j As Integer
    
    With ActiveSheet
        If Rng Is Nothing Then
            If Selection.Cells.Count = 1 Then
                Set Rng = Range(Rows(1), _
                    Rows(.Cells.Find("*", SearchOrder:=xlByRows, LookIn:=xlValues, SearchDirection:=xlPrevious).Row))
            ElseIf Selection.Rows.Count = 1 Then
                If Application.CountA(Selection.EntireRow) = 0 Then Selection.EntireRow.Delete
                    Exit Sub
            Else
                Set Rng = Selection
            End If
        End If
        For i = Rng.Rows.Count To 1 Step -1
            For j = Rng.Columns.Count To 1 Step -1
                If Rng.Cells(i, j).Value = sValue Then Rng.Rows(i).EntireRow.Delete
            Next j
        Next i
    End With
End Sub




'DeleteColumnCriteria
'Delete columns with a certain value in them
'
'
Sub DelColCrit(sValue As String, Optional Rng As Range)
    Dim i As Integer
    Dim j As Integer
    
    With ActiveSheet
        If Rng Is Nothing Then
            If Selection.Cells.Count = 1 Then
                Set Rng = Range(Rows(1), _
                    Rows(.Cells.Find("*", SearchOrder:=xlByRows, LookIn:=xlValues, SearchDirection:=xlPrevious).Row))
            ElseIf Selection.Rows.Count = 1 Then
                If Application.CountA(Selection.EntireRow) = 0 Then Selection.EntireRow.Delete
                    Exit Sub
            Else
                Set Rng = Selection
            End If
        End If
        For i = Rng.Rows.Count To 1 Step -1
            For j = Rng.Columns.Count To 1 Step -1
                If Rng.Cells(i, j).Value = sValue Then Rng.Columns(j).EntireColumn.Delete
            Next j
        Next i
    End With
End Sub




'IsWorkbookOpen
'Will try to find a workbook with a similar name than that which is given. It
'works by finding a substring so that the function works the same way for worsheet with
'variable name (e.g. BalanceSheet-2009, BalanceSheet-2010 etc... ) If the worksheet is not found
'will return a nullstring.
'
'Usage
'Usually use it at the begining if data is neede from a certain sheet and the user needs to open
'it before hand  - This might be simpler if the location of the file is not known. It is easier 
'to make the user open it manually
'
'If sIsWorkbookOpen("BalanceSheet") = vbNullString then
'	MsgBox "Error - Balancesheet has to be open", vbOKOnly
'	exit sub
'end if
''.... continue work normally
'

Function IsWorkbookOpen(sWorkBookName As String) As String
    Dim OpenWorkbook As Workbook
    IsWorkbookOpen = vbNullString
    For Each OpenWorkbook In Application.Workbooks
        If InStr(OpenWorkbook.Name, sWorkBookName) <> 0 Then
            IsWorkbookOpen = OpenWorkbook.Name
            Exit Function
        End If
    Next OpenWorkbook
End Function




'HideSheet
'Will make the specified or active sheet Very Hidden
'which can only be revealed with VB
'
'Usage
'HideSheet("Balance Sheet")
'

Sub HideSheet(Optional SheetName as String)
    If SheetName = "" Then SheetName = ActiveSheet.Name
    Sheets(SheetName).Visible = xlSheetVeryHidden
End Sub



'ShowAllHiddenSheets
'Will make all  sheets in the active workbook visible
'
'Usage
'ShowAllHiddenSheets
'

Sub ShowAllHiddenSheets()
    Dim i As Integer
    For i = 1 To ActiveWorkbook.Sheets.Count
        Sheets(i).Visible = xlSheetVisible
    Next i
End Sub



'DeleteEmptySheets
'Delete all empty sheets from wither the specified workbook
'or the active workbook if nothing is specified
'
'Usage
'DeleteEmpthSheets("ThisWorkbook.xls")
'

Sub DeleteEmptySheets( _
    Optional wbName As String)
    Dim Ws As Worksheet
    If wbName = "" Then wbName = ActiveWorkbook.Name
    For Each Ws In Workbooks(wbName).Worksheets
        If Application.WorksheetFunction.CountA(Ws.UsedRange.Cells) = 0 Then Ws.Delete
    Next Ws
End Sub



'AddCustomList
'Adds a custom List to excel. This is usefull use the excel sort function for custom
'lists (e.g. sorting by credit rating AAA+, AAA, AAA-, .....)
'
'Usage
'
'

Function AddCustomList(ByRef sCustomList As Variant, bAddList As Boolean) As Variant
    
    Dim lLoop As Long
    Dim iCustomLists As Integer
    
    Dim i As Integer
    
    Dim sTempCustList()
 
    For lLoop = 1 To Application.CustomListCount
        sTempCustList = Application.GetCustomListContents(lLoop)
            
        If UBound(sTempCustList) = UBound(sCustomList) Then
            For i = 1 To UBound(sCustomList)
                If sTempCustList(i) <> sCustomList(i) Then
                    GoTo NextList
                End If
            Next i
            If i = UBound(sCustomList) + 1 Then GoTo ListPresent
        End If
NextList:
    Next lLoop
           
    If bAddList Then
        Application.AddCustomList (sCustomList)
        AddCustomList = lLoop + 1
    Else
        AddCustomList = vbNullString
    End If
        
    Exit Function
ListPresent:
    AddCustomList = lLoop + 1
End Function



'WRSToOverride
'Change delimitors from custom ones to normal vb to use in code
'
'
'
'
Function WRSToOverride(ByVal sNumber As String) As String
    Dim sWRS As String, sWRSThousand As String, sWRSDecimal As String
    Dim sXLThousand As String, sXLDecimal As String
    
    If Val(Application.Version) >= 10 Then
        If Not Application.UseSystemSeparators Then
            sWRS = Format(1000, "#,##0.00")
            sWRSThousand = Mid(sWRS, 2, 1)
            sWRSDecimal = Mid(sWRS, 6, 1)
            sXLThousand = Application.ThousandsSeparator
            sXLDecimal = Application.DecimalSeparator
            sNumber = Replace(sNumber, sWRSThousand, vbTab)
            sNumber = Replace(sNumber, sWRSDecimal, sXLDecimal)
            sNumber = Replace(sNumber, vbTab, sXLThousand)
        End If
    End If
    WRSToOverride = sNumber
End Function




'ForceText
'
'
'
'
'

Sub ForceText(Col As Integer)
    Dim i As Integer
    Dim j As String
        For i = 1 To LastRow
            j = Cells(i, Col + 1).Value
            Cells(i, Col).Formula = j
        Next i
End Sub




'CopyUniqueValues
'Copies the rows of all the unique values from a column to another
'sheet.
'

'

Sub CopyUniqueValues( _
    shInput As String, _
    shOutput As String, _
    clInput As String, _
    clOutput As String)
    Dim i As Integer, j As Integer
    
    j = 1
    
    For i = 1 To LastRow(shInput)
        If Application.WorksheetFunction.CountIf(Sheets(shOutput).Columns(clOutput), Sheets(shInput).Cells(i, clInput)) = 0 Then
            Rows(i).Copy
            Application.Paste Destination:=Sheets(shOutput).Rows(j)
            j = j + 1
        End If
    Next i
    
End Sub


'
'
'
'
'
'

Sub TextToWordArray(sText As String, sWordSeparator As String, TempArr())

    Dim i As Integer
    Dim iPos As Long
    Dim iNextPos As Long

    sText = Replace(sText, Chr(10), " ")

    If InStr(sText, sWordSeparator) = 0 Then GoTo Only1Word
        
    i = 0
    iPos = InStr(1, sText, sWordSeparator)
    Do
        i = i + 1
        iPos = InStr(iPos + 1, sText, sWordSeparator)
    Loop Until iPos = 0

    ReDim TempArr(0 To i)
    
    i = 0
    
    iPos = 1
    iNextPos = InStr(iPos + 1, sText, sWordSeparator)
    
    Do
        TempArr(i) = Trim(Mid(sText, iPos, iNextPos - iPos))
        i = i + 1
        iPos = iNextPos + 1
        iNextPos = InStr(iPos + 1, sText, sWordSeparator)
    Loop Until iNextPos = 0
    
    TempArr(i) = Trim(Right$(sText, (Len(sText) - iPos)))


    ReDim Preserve TempArr(LBound(TempArr()) To UBound(TempArr()))

    Exit Sub

Only1Word:
    ReDim TempArr(0 To 0)
    TempArr(0) = sText
    
End Sub






'BrowseFile
'Shows the file browser dialog and returns the file path or vbNullString if nothing is selected
'
'Usage
'fileToOpen = BrowseFile
'
'
Function BrowseFile() As String
    Dim fPath
    fPath = Application.GetOpenFilename()
    If fPath <> False Then ; BrowseFile = fPath;    Else;   BrowseFile = vbNullString
End Function



'BrowseForFolder
'Shows the folder browser dialog and returns the file path or vbNullString if nothing is selected
'
'Usage
'fileToOpen = BrowseFile
'
'

Function BrowseForFolder( _
    Optional OpenAt As Variant) As Variant
    'Function purpose:  To Browser for a user selected folder.
    'If the "OpenAt" path is provided, open the browser at that directory
    'NOTE:  If invalid, it will open at the Desktop level
    Dim ShellApp As Object

    'Create a file browser window at the default folder
    Set ShellApp = CreateObject("Shell.Application"). _
            BrowseForFolder(0, "Please choose a folder", 0, OpenAt)
            
    'Set the folder to that selected.  (On error in case cancelled)
    On Error Resume Next
    BrowseForFolder = ShellApp.self.path
    On Error GoTo 0

    'Destroy the Shell Application
    Set ShellApp = Nothing

    'Check for invalid or non-entries and send to the Invalid error
    'handler if found
    'Valid selections can begin L: (where L is a letter) or
    '\\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(BrowseForFolder, 2, 1)
    Case Is = ":"
        If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
    Case Is = "\"
        If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
    Case Else
        GoTo Invalid
    End Select
    Exit Function
    
    Invalid:
    'If it was determined that the selection was invalid, set to False
    BrowseForFolder = False
End Function



'DefaultPrintSetup
'
'
'
'
'
'
'
'
'
Sub DefaultPrintSetup(PaperType As Excel.XlPaperSize, PaperDir As Excel.XlPageOrientation, Header As Boolean)
    With ActiveSheet.PageSetup
        If Header = True Then
            .LeftHeader = "&F&A"
            .CenterHeader = ""
            .RightHeader = "Page &P de &N"
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = "&Z&F"
        End If
        .LeftMargin = Application.InchesToPoints(0.78740157480315)
        .RightMargin = Application.InchesToPoints(0.78740157480315)
        .TopMargin = Application.InchesToPoints(0.984251968503937)
        .BottomMargin = Application.InchesToPoints(0.984251968503937)
        .HeaderMargin = Application.InchesToPoints(0.511811023622047)
        .FooterMargin = Application.InchesToPoints(0.511811023622047)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = PaperDir
        .Draft = False
        .PaperSize = PaperType
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .PrintErrors = xlPrintErrorsDisplayed
    End With
End Sub







Function compareRanges( _
    rng1 As Excel.Range, _
     rng2 As Excel.Range, _
     Optional compareMin As Boolean = False, _
     Optional transpose As Boolean = False) As Boolean
    Dim i As Integer, j As Integer
    Dim rwMax As Integer, clMax As Integer
    
    If Not compareMin Then
        If rng1.Rows.Count <> rng2.Rows.Count Or rng1.Columns.Count <> rng2.Columns.Count Then
            compareRanges = False
        End If
        rwMax = rng1.Rows.Count
        clMax = rng2.Columns.Count
    Else    
        rwMax = Application.WorksheetFunction.Min(rng1.Rows.Count, rng2.Rows.Count)
        rwMax = Application.WorksheetFunction.Min(rng1.Columns.Count, rng2.Columns.Count)
    End If
    
    If Not transpose Then
        For i = 1 To rwMax
            For j = 1 To clMax
                If rng1.Cells(i, j).Value <> rng2.Cells(i, j).Value Then
                    compareRanges = False
                    Exit Function
                End If
            Next j
        Next i
        compareRanges = True
    Else
        For i = 1 To rwMax
            For j = 1 To clMax
                If rng1.Cells(i, j).Value <> rng2.Cells(j, i).Value Then
                    compareRanges = False
                    Exit Function
                End If
            Next j
        Next i
     compareRanges = True
    End If
    End Function




























