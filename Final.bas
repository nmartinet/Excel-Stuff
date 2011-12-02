Option Explicit
Option Base 0
'Helpers
'v 11.11.1
'
'ChangeLog
'
'
'
'
'
'
'


'**********************************************************************************************'
'*
'*  Internal Subs
'*
'**********************************************************************************************'
Private Sub aa_Internal_Subs_Start()
End Sub

Sub DisableUpdatesErrors(Disabled As Boolean, Optional StatusBarMessage As String = "Working...")
  Static calcType As Long, refType As Long
  
  On Error Resume Next
  With Excel.Application
    If Disabled = True Then
      calcType = Application.Calculation
      refType = Application.ReferenceStyle
      
      .Cursor = xlWait
      .StatusBar = StatusBarMessage
      .EnableCancelKey = xlInterrupt
      .ScreenUpdating = False
      .EnableEvents = False
      .DisplayAlerts = False
      .Calculation = xlCalculationManual
      .ReferenceStyle = xlR1C1
    Else
      .Cursor = xlDefault
      .StatusBar = False
      .EnableCancelKey = xlInterrupt
      .ScreenUpdating = True
      .EnableEvents = True
           
      If Not calcType = 0 Then
        .Calculation = calcType
        .ReferenceStyle = refType
        calcType = 0
        refType = 0
      End If
      
    End If
  End With
End Sub

Sub TextToWordArray(sText As String, sWordSeparator As String, TempArr())
  Dim i As Integer, iPos As Long, iNextPos As Long
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

'**********************************************************************************************'
'*
'*  Internal Functions
'*
'**********************************************************************************************'
Private Sub aa_Internal_Functions_Start()
End Sub

Function LastRow(Optional ws As Worksheet)
  On Error Resume Next
  If ws Is Nothing Then Set ws = ActiveSheet
  LastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, LookIn:=xlValues, SearchDirection:=xlPrevious).EntireRow.Row
  If LastRow = "" Then LastRow = 0
End Function

Function LastColumn(Optional ws As Worksheet)
  On Error Resume Next
  If ws Is Nothing Then Set ws = ActiveSheet
  LastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, LookIn:=xlValues, SearchDirection:=xlPrevious).EntireColumn.Column
  If LastColumn = "" Then LastColumn = 0
End Function

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

Function WRSToOverride(ByVal sNumber As String) As String
  Dim sWRS As String, sWRSThousand As String, sWRSDecimal As String
  Dim sXLThousand As String, sXLDecimal As String
  
  If val(Application.Version) >= 10 Then
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

Function BrowseFile() As String
  Dim fPath
  fPath = Application.GetOpenFilename()
  If fPath <> False Then: BrowseFile = fPath: Else: BrowseFile = vbNullString: End If
End Function

Function BrowseForFolder(Optional OpenAt As Variant) As Variant
  Dim ShellApp As Object
  
  Set ShellApp = CreateObject("Shell.Application").BrowseForFolder(0, "Please choose a folder", 0, OpenAt)
  
  On Error Resume Next
  BrowseForFolder = ShellApp.self.Path
  On Error GoTo 0
  
  Set ShellApp = Nothing
  
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
  BrowseForFolder = False
End Function

'**********************************************************************************************'
'*
'*  Reference Funtions
'*
'**********************************************************************************************'
Private Sub aa_Reference_Functions_Start()
End Sub

Function StringToRefStyle(s As String)
  Dim tmp As Stringm, fromRng As String
  
  'try an convert ref to A1 - Range object
  'cannot be set from R1C1 ref
  tmp = s
  On Error Resume Next
  tmp = Application.ConvertFormula(s, xlR1C1, xlA1, True, ActiveCell)
  
  If InStr(1, tmp, "[") <> 0 Then
    tmp = Mid(t, 1, Len(tmp) - 1)
    tmp = Replace(tmp, "!'", "!")
  Else
    tmp = Replace(tmp, "'", "")
  End If
  
  'try to assign a range from the string
  fromRng = Range(tmp).Address(False, False, xlA1, False, ActiveCell)
  On Error GoTo 0
  
  If fromRng = vbNullString Then
    StringToRefStyle = False
  Else
    tmp = vbNullString
    On Error Resume Next
    tmp = Application.ConvertFormula(s, xlA1, xlR1C1, False, ActiveCell)
    On Error GoTo 0
    'if t converted - A1 style
    If tmp <> vbNullString And tmp <> s Then
      StringToRefStyle = "A1"
    Else
      tmp = vbNullString
      On Error Resume Next
      tmp = Application.ConvertFormula(s, xlR1C1, xlA1, False, ActiveCell)
      On Error GoTo 0
      If tmp = s Then
        'not conveted but could apply ganre - name
        StringToRefStyle = "NAME"
      ElseIf tmp <> vbNullString Then
        'converted - xlR1C1
        StringToRefStyle = "R1C1"
      End If
    End If
  End If
End Function

Function GenerateReference(str As String, absolute As Boolean, Optional refType As XlReferenceStyle = xlR1C1) As String
  Dim t As String
  
  'if str is between double quotes - string literal
  If Left(str, 1) = """" And Right(str, 1) = """" Then
    GenerateReference = str
    Exit Function
  End If
  
  t = StringToRefStyle(str)
  
  If t = "A1" Then
    GenerateReference = Range(str).Address(absolute, absolute, refType, False, ActiveCell)
  ElseIf t = "R1C1" Then
    GenerateReference = Range(Application.ConvertFormula(str, xlR1C1, xlA1, True, ActiveCell)).Address(absolute, absolute, refType, False, ActiveCell)
  ElseIf t = "NAME" Then
    GenerateReference = str
  ElseIf t = False Then
    GenerateReference = """" & str & """"
  End If
  
End Function

Function GenerateRefRC(rng As Range, Optional getRow As Boolean = True, Optional getSheetName As Boolean = False)
  Dim refStr As String
  If getRow Then
    refStr = "R" & rng.Rows(1).Row
  Else
    refStr = "C" & rng.Columns(1).Column
  End If
  
  If getSheetName Then
    GenerateRefRC = "'" & rng.Parent.Name & "'!" & refS
  Else
    GenerateRefRC = refStr
  End If
End Function


'**********************************************************************************************'
'*
'*  Data Manipulation
'*
'**********************************************************************************************'
Private Sub aa_Data_Manipulation_Start()
End Sub


Sub ClearData(ws As Worksheet, keyRow As Integer, keyCol As Integer, targetValue As String, keepFiltered As Boolean)
  Dim startRow As Long, endRow As Long, lRow As Long
  Dim rng As Range
  
  With ws
    On Error Resume Next
    .AutoFilter.Sort.SortFields.Clear
    .ShowAllData
    On Error GoTo 0
    Set rng = .Cells(keyRow, keyCol)
    With .AutoFilter.Sort
      .SortFields.Clear
      .SortFields.Add Key:=rng, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
    End With
    
    On Error Resume Next
    startRow = .Columns(keyCol).Find(What:=targetValue, After:=.Cells(keyRow, keyCol), LookIn:=xlValues, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Row
    On Error GoTo 0
    
    If startRow = 0 Then Exit Sub
    
    endRow = .Columns(keyCol).Find(What:=targetValue, After:=.Cells(keyRow, keyCol), LookIn:=xlValues, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False).Row
    
    If keepFiltered Then
      lRow = LastRow(ws)
      If endRow <> lRow Then .Range(.Rows(endRow + 1), .Rows(lRow)).Delete
      If startRow <> keyRow + 1 Then .Range(.Rows(2), .Rows(startRow - 1)).Delete
    Else
      .Range(.Rows(startRow), .Rows(endRow)).Delete
    End If
  End With
End Sub

Sub CopyData(ws As Worksheet, keyRow As Integer, keyCol As Integer, targetValue As String)
  Dim startRow As Long, endRow As Long, lRow As Long
  Dim rng As Range
  
  With ws
    On Error Resume Next
    .AutoFilter.Sort.SortFields.Clear
    .ShowAllData
    On Error GoTo 0
    Set rng = .Cells(keyRow, keyCol)
    With .AutoFilter.Sort
      .SortFields.Clear
      .SortFields.Add Key:=rng, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
      .Header = xlYes
      .MatchCase = False
      .Orientation = xlTopToBottom
      .SortMethod = xlPinYin
      .Apply
    End With
    
    On Error Resume Next
    
    startRow = .Columns(keyCol).Find(What:=targetValue, After:=.Cells(keyRow, keyCol), LookIn:=xlValues, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Row
    
    On Error GoTo 0
    
    If startRow = 0 Then Exit Sub
    
    endRow = .Columns(keyCol).Find(What:=targetValue, After:=.Cells(keyRow, keyCol), LookIn:=xlValues, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False).Row
    
    .Range(.Rows(startRow), .Rows(endRow)).Copy
  End With
End Sub

Private Sub DeleteEmptyRows()
  Dim rng As Range, i As Long
  On Error GoTo Err_Handl
  With ActiveSheet
    If Selection.Cells.Count = 1 Then
      Set rng = Range(Rows(1), Rows(.Cells.Find("*", SearchOrder:=xlByRows, LookIn:=xlValues, SearchDirection:=xlPrevious).Row))
    ElseIf Selection.Rows.Count = 1 Then
      If Application.CountA(Selection.EntireRow) = 0 Then Selection.EntireRow.Delete
      Exit Sub
    Else
      Set rng = Selection
    End If
    For i = rng.Rows.Count To 1 Step -1
      If Application.CountA(rng.Rows(i).EntireRow) = 0 Then rng.Rows(i).EntireRow.Delete
    Next i
  End With
  Exit Sub
Err_Handl:
  MsgBox "Error - Sheet might contain only empty rows", vbOKOnly
End Sub

Private Sub DeleteEmptyColumns()
  Dim rng As Range, i As Long
  On Error GoTo Err_Handl
  With ActiveSheet
    If Selection.Cells.Count = 1 Then
      Set rng = Range(Columns(1), Columns(.Cells.Find("*", SearchOrder:=xlByColumns, LookIn:=xlValues, SearchDirection:=xlPrevious).Column))
    ElseIf Selection.Columns.Count = 1 Then
      If Application.CountA(Selection.EntireColumn) = 0 Then Selection.EntireColumn.Delete
      Exit Sub
    Else
      Set rng = Selection
    End If
    For i = rng.Columns.Count To 1 Step -1
      If Application.CountA(rng.Columns(i).EntireColumn) = 0 Then rng.Columns(i).EntireColumn.Delete
    Next i
  End With
  Exit Sub
Err_Handl:
  MsgBox "Error - Sheet might contain only empty columns", vbOKOnly
End Sub

Private Sub DelRowCrit(sValue As String, Optional rng As Range)
  Dim i As Integer, j As Integer
  
  With ActiveSheet
    If rng Is Nothing Then
      If Selection.Cells.Count = 1 Then
        Set rng = Range(Rows(1), Rows(.Cells.Find("*", SearchOrder:=xlByRows, LookIn:=xlValues, SearchDirection:=xlPrevious).Row))
      ElseIf Selection.Rows.Count = 1 Then
        If Application.CountA(Selection.EntireRow) = 0 Then Selection.EntireRow.Delete
        Exit Sub
      Else
        Set rng = Selection
      End If
    End If
    For i = rng.Rows.Count To 1 Step -1
      For j = rng.Columns.Count To 1 Step -1
        If rng.Cells(i, j).Value = sValue Then rng.Rows(i).EntireRow.Delete
      Next j
    Next i
  End With
End Sub

Private Sub DelColCrit(sValue As String, Optional rng As Range)
  Dim i As Integer, j As Integer
  
  With ActiveSheet
    If rng Is Nothing Then
      If Selection.Cells.Count = 1 Then
        Set rng = Range(Rows(1), Rows(.Cells.Find("*", SearchOrder:=xlByRows, LookIn:=xlValues, SearchDirection:=xlPrevious).Row))
      ElseIf Selection.Rows.Count = 1 Then
        If Application.CountA(Selection.EntireRow) = 0 Then Selection.EntireRow.Delete
        Exit Sub
      Else
        Set rng = Selection
      End If
    End If
    For i = rng.Rows.Count To 1 Step -1
      For j = rng.Columns.Count To 1 Step -1
        If rng.Cells(i, j).Value = sValue Then rng.Columns(j).EntireColumn.Delete
      Next j
    Next i
  End With
End Sub


'**********************************************************************************************'
'*
'*  Queries
'*
'**********************************************************************************************'
Private Sub aa_Queries_Start()
End Sub


Sub ExecuteQuery(q As String, out As Range, Optional headers = True)
  Dim cn As Object, rs As Object
  Dim wbFullName As String, i As Integer
  
  wbFullName = CreateTempWb(ActiveWorkbook)
  
  Set cn = CreateObject("ADODB.Connection")
  cn.Open GenerateQueryString(wbFullName)
  
  Set rs = CreateObject("ADODB.Recordset")
  rs.Open q, cn
  
  If Not (rs.BOF Or rs.EOF) = 0 Then
    out = "NO_DATA"
  Else
    If headers Then
      For i = 0 To rs.Fields.Count - 1
        out.Offset(0, i).Value = rs.Fields(i).Name
      Next
      out.Offset(1, 0).CopyFromRecordset rs
    Else
      out.CopyFromRecordset rs
    End If
  End If
   
  rs.Close:           cn.Close
  Set rs = Nothing:   Set cn = Nothing
  Kill wbFullName
  
End Sub

Sub ExecuteQueries(q, out, postQ)
  Dim cn As Object, rs As Object
  Dim i As Integer, j As Integer
  Dim v As Variant
  Dim wbFullName As String
  
  wbFullName = CreateTempWb(ActiveWorkbook)
  
  Set cn = CreateObject("ADODB.Connection")
  cn.Open GenerateQueryString(wbFullName)
  
  Set rs = CreateObject("ADODB.Recordset")
    
  For i = 0 To UBound(q)
  
    rs.Open q(i), cn
    If Not (rs.BOF Or rs.EOF) = 0 Then
       out(i) = "NO_DATA"
    Else
      If True Then
        For j = 0 To rs.Fields.Count - 1
          out(i).Offset(0, j).Value = rs.Fields(j).Name
        Next
        out(i).Offset(1, 0).CopyFromRecordset rs
      Else
        out(i).CopyFromRecordset rs
      End If
    End If
    rs.Close
    
    If Not postQ(i) = vbNullString Then
      Run (postQ(i))
    End If
    
  Next
  
  cn.Close
  Set rs = Nothing:  Set cn = Nothing
  Kill wbFullName
  
End Sub

Function CreateTempWb(wb As Workbook) As String
  Dim wbFullName As String
  wbFullName = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".xlsm"
  wb.SaveCopyAs (wbFullName)
  CreateTempWb = wbFullName
End Function

Function GenerateQueryString(wbName As String, Optional connType As String = "Excel") As String
  GenerateQueryString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & _
        wbName & ";Extended Properties=""Excel 12.0 Macro;HDR=YES;IMEX=1"";"
End Function


'**********************************************************************************************'
'*
'*  Worksheet Formulas
'*
'**********************************************************************************************'
Private Sub aa_Worksheet_Formulas_Start()
End Sub


Function vLookUpConcat(LookUpValue As String, LookupRange As Range, ValueCol As Integer, Optional Separator As String = "; ")
  Dim s As String
  Dim i As Variant

  For i = 1 To LookupRange.Rows.Count
    If LookupRange.Cells(i, 1).Value = LookUpValue Then s = s & LookupRange.Cells(i, ValueCol).Value & Separator
  Next
  
  If Len(s) > 0 Then s = Left(s, Len(s) - Len(Separator))
  vLookUpConcat = s
  
End Function

Function vLookUpConcatUnique(LookUpValue As String, LookupRange As Range, ValueCol As Integer, Optional Separator As String = "; ")
  Dim s As String
  Dim i As Long
  Dim col As New Collection

  For i = 1 To LookupRange.Rows.Count
    If LookupRange.Cells(i, 1).Value = LookUpValue Then
      On Error Resume Next
      col.Add LookupRange.Cells(i, ValueCol).Value, LookupRange.Cells(i, ValueCol).Value
      On Error GoTo 0
    End If
  Next
  
  If col.Count > 0 Then
    For i = 1 To col.Count - 1
      s = s & col.Item(i) & Separator
    Next
    s = s & col.Item(col.Count)
  End If
  
  vLookUpConcatUnique = s
    
End Function

'**********************************************************************************************'
'*
'*  End user Subs
'*
'**********************************************************************************************'
Private Sub aa_end_User_Subs_Start()
End Sub

Sub HideSheet(Optional ws As Worksheet)
  If ws Is Nothing Then Set ws = ActiveSheet
  ws.Visible = xlSheetVeryHidden
End Sub

Private Sub UnhideAlll()
  Dim ws As Worksheet
  For Each ws In ActiveWorkbook.Sheets
    ws.Visible = xlSheetVisible
  Next
End Sub

Private Sub LockAll()
  Dim ws As Worksheet
  Dim pw As String
  Dim pw2 As String
  
  pw = InputBox("Password", "Password", "")
  pw2 = InputBox("Repeat Password", "Repeat Password", "")
  
  If pw <> pw2 Then
    MsgBox "Password did not match", vbOKOnly
    Exit Sub
  End If
  
  Application.ScreenUpdating = False
  For Each ws In ActiveWorkbook.Sheets
    On Error Resume Next
    ws.Protect pw
    On Error GoTo 0
  Next
  Application.ScreenUpdating = True
    
End Sub

Private Sub UnlockAll()
  Dim ws As Worksheet, pw As String

  pw = InputBox("Password", "Password", "")
  
  Application.ScreenUpdating = False
  For Each ws In ActiveWorkbook.Sheets
    On Error Resume Next
    ws.Unprotect pw
    On Error GoTo 0
  Next
  Application.ScreenUpdating = True
    
End Sub

Sub DeleteEmptySheets(Optional wb As Workbook)
  Dim ws As Worksheet
  If wb Is Nothing Then Set wbName = ActiveWorkbook
  For Each ws In wb.Worksheets
    If Application.WorksheetFunction.CountA(ws.UsedRange.Cells) = 0 Then ws.Delete
  Next ws
End Sub

Sub CopyUniqueValues(shInput As String, shOutput As String, clInput As String, clOutput As String)
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

Function compareRanges(rng1 As Excel.Range, rng2 As Excel.Range, Optional compareMin As Boolean = False, _
                        Optional transpose As Boolean = False) As Boolean
  Dim i As Integer, j As Integer
  Dim rwMax As Integer, clMax As Integer
  
  If Not compareMin Then
    If rng1.Rows.Count <> rng2.Rows.Count Or rng1.Columns.Count <> rng2.Columns.Count Then compareRanges = False
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


'**********************************************************************************************'
'*
'*  Name & List Manipulation
'*
'**********************************************************************************************'
Private Sub aa_Name_and_List_Manipulation_Start()
End Sub

Function AddCustomList(ByRef sCustomList As Variant, bAddList As Boolean) As Variant
  Dim lLoop As Long, iCustomLists As Integer, i As Integer
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

Private Sub AddDynamicNamePrompt()
  Dim n As String
  Dim v As Variant
  
  n = InputBox("Range Name", "Range Name", "")
  
  Application.ScreenUpdating = False
  For Each v In ActiveWorkbook.Names
    If n = v.Name Then
      Application.ScreenUpdating = True
      MsgBox "Name already exists", vbOKOnly, "Name already exists"
      Exit Sub
    End If
  Next v
  
  AddDynamicName n, ActiveCell
  
End Sub

Sub AddDynamicName(n As String, rng As Range)
  ActiveWorkbook.Names.Add Name:=n, _
                           RefersToR1C1:="=OFFSET(" & rng.Address(True, True, xlR1C1, False, ActiveCell) & _
                              ",0,0,COUNTA(" & GenerateRefRC(ActiveCell, False) & _
                                "), COUNTA(" & GenerateRefRC(ActiveCell, True) & "))"
End Sub


'**********************************************************************************************'
'*
'*  Generate Mail Subs
'*
'**********************************************************************************************'
Private Sub aa_Generate_Mail_Subs_Start()
End Sub

Sub GenHTMLMail(rng As Range, toEmail As String, subject As String, Optional cc As String = vbNullString, _
                          Optional bcc As String = vbNullString)
  'adapted from  http://msdn.microsoft.com/en-us/library/ff519602(office.11).aspx
  Dim OutlookApp As Object, email As Object
    
  'create outlook and email objects
  Set OutlookApp = CreateObject("Outlook.Application")
  Set email = OutlookApp.createitem(0)
  
  'set email parameters
  'can chage .Display to .Send automatically send the email instead of display
  'On Error Resume Next
  With email
    .To = toEmail
    .cc = cc
    .bcc = bcc
    .subject = subject
    .HTMLBody = RangetoHTML(rng)
    .Display
  End With
  On Error GoTo 0
  
  Set email = Nothing
  Set OutlookApp = Nothing
  
End Sub

Function RangetoHTML(rng As Range)
  Dim fso As Object, ts As Object
  Dim TempFile As String, TempWB As Workbook
  Dim c As Variant
  Dim origSize As Integer, origFont As String
    
  TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
  
  ' Copy the range and create a workbook to receive the data.
  rng.Copy
  Set TempWB = Workbooks.Add(1)
  
  With TempWB.Sheets(1)
  
    .Cells(1).PasteSpecial Paste:=8
    .Cells(1).PasteSpecial xlPasteValues, , False, False
    .Cells(1).PasteSpecial xlPasteFormats, , False, False
    .Cells.Columns.AutoFit
    'add back links
    For Each c In .UsedRange
      If (Left(c.Value, 2) = "\\") Or (Left(c.Value, 7) = "http:\\") Then
        origSize = c.Font.Size
        origFont = c.Font.Name
        .Hyperlinks.Add anchor:=c, Address:=c.Value, TextToDisplay:=c.Value
        With c.Font
          .Size = origSize
          .Name = origFont
          .Underline = xlUnderlineStyleSingle
          .ThemeColor = xlThemeColorHyperlink
        End With
      End If
    Next c
    
    Application.CutCopyMode = False
    On Error Resume Next
    .DrawingObjects.Visible = True
    .DrawingObjects.Delete
    On Error GoTo 0
    
  End With
  
  'Publish the sheet to an .htm file.
  With TempWB.PublishObjects.Add(xlSourceRange, _
    TempFile, TempWB.Sheets(1).Name, _
        TempWB.Sheets(1).UsedRange.Address(ReferenceStyle:=Application.ReferenceStyle), _
        xlHtmlStatic, "RSDE - Suivi", "")
    .Publish (True)
    .AutoRepublish = False
  End With
  
  'Read all data from the .htm file into the RangetoHTML subroutine.
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
  RangetoHTML = ts.ReadAll
  ts.Close
  RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
  "align=left x:publishsource=")
  
  'Close TempWB.
  TempWB.Close savechanges:=False
  
  'Delete the htm file.
  Kill TempFile
  
  Set ts = Nothing
  Set fso = Nothing
  Set TempWB = Nothing
End Function


'**********************************************************************************************'
'*
'*  Array functions
'*
'**********************************************************************************************'
Private Sub aa_Array_Functions_Start()
End Sub

Sub Push(ByRef arr, ByVal val)
  Dim tmp()
  tmp = arr
  ReDim Preserve tmp(0 To UBound(arr) + 1)
  tmp(UBound(arr) + 1) = val
  arr = tmp
End Sub

Function Pop(ByRef arr)
  Dim tmp(), i
  ReDim tmp(0 To UBound(arr) - 1)
  For i = 0 To UBound(tmp)
    tmp(i) = arr(i)
  Next
  Pop = arr(UBound(arr))
  arr = tmp
End Function




