Attribute VB_Name = "mActionQueries"
Option Explicit
Option Private Module
Public STATUS As String           'static variable for err hndl
Public EXECUTED_QUERIES As String 'static variable - holds the successfully executed queries
                                  'and the number of affected entries
                  
'-----------------------------------------------------------
'ADD REFERENCE TO DAO360.DLL - MICROSOFT DAO 3.6 OBJECT LIBRARY
'if it is not in the references list, it is located in
' C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll
'
'last update - 2009-06-08 - 11:12

Sub runActionQueriesSet(dbSourcePath As String, _
                        dbSourceType As String, _
                        dbTargetPath As String, _
                        dbTargetType As String, _
                        Queries() As String, _
                        Optional dbUser = "Admin", _
                        Optional dbPassword As String)

'-----------------------------------------------------------------------------------'
'Sub runActionQueriesSet
'Sub will Execute a loop of queries on the target database
'since this uses DAO, it is essentially used for Jet databases
'This is only for queries that will not return anything ie action queries
'such as INSERT INTO, DELETE, UPDATE, etc.
'See the 2 public variables after the sub ran for statuses
'-----------------------------------------------------------------------------------'

  Dim wrkSpace  As DAO.Workspace
  Dim dbSource  As DAO.Database   'Excel database
  Dim dbTarget  As DAO.Database   'Access database
  Dim i         As Integer        'loop in query
  Dim dErr      As DAO.Error
  Dim failed    As Boolean        'if current query fails -> true and not
                                  'added to EXECUTED_QUERIES
  Dim rollB     As Boolean        'if at least 1 query fails, will be true
                                  'and all changes will be rolled back
  
  
  'goto error reporting only for connection errors
  'most common - wrong dbPath or wrong User/Password
  On Error GoTo connectErr
    'create workspace and link to dbs
    Set wrkSpace = CreateWorkspace("", dbUser, dbPassword)
    Set dbSource = OpenDatabase(dbSourcePath, False, True, dbSourceType)
    Set dbTarget = wrkSpace.OpenDatabase(dbTargetPath, False, False, dbTargetType)
  On Error GoTo 0
  
  'init public consts
  EXECUTED_QUERIES = ""
  STATUS = ""
  
  'begin transaction to rollback in case of error
  wrkSpace.BeginTrans
  On Error GoTo ErrHndl
    
    'loop through all the queries
    For i = 0 To UBound(Queries)
      'reset failed to false
      failed = False

      'try to execute query
      dbTarget.Execute Queries(i), dbFailOnError
      
      'if query didn't fail - append it to EXECUTED_QUERIES with it's no. of rows affected
      If Not (failed) Then
        EXECUTED_QUERIES = EXECUTED_QUERIES & Queries(i) & _
                          "Records Affected: " & dbTarget.RecordsAffected & vbCrLf
      End If
    Next i
  
  'if rollB is true, then at
  'least one query failed -> go to handler
  If rollB Then GoTo rBack
  
  'if no queries failed, commit the changes
  wrkSpace.CommitTrans
  
  'in debug - rollback
  'uncomment committrans in prod
  'wrkSpace.Rollback
  
  On Error GoTo 0

  'free memory, close connections
  Set dbSource = Nothing
  Set dbTarget = Nothing
  Set wrkSpace = Nothing
  
  'no problems occured -> status = ok
  STATUS = "OK - All queries Executed Succesfully"

Exit Sub

'ERROR HANDLING
connectErr:
  'in case of connection errors
  'err.num 3024 -> wrong db path/name
  'err.num 3028 -> wrong user name/password
  'other num ???
  If Err.Number = 3024 Then
    MsgBox "Base de Données n'existe pas sous: " & vbCrLf & _
            dbTargetPath & vbCrLf & _
            dbSourcePath, vbCritical + vbOKOnly
  ElseIf Err.Number = 3028 Then
    MsgBox "Nom d'utilisateur ou mot de passe incorrecte" _
            , vbCritical + vbOKOnly
  Else
    MsgBox "Erreur inconnue", vbCritical + vbOKOnly
  End If
  
  'since couldn't connect to at least one DB
  'no queries were executed - end completely - no need for log
  End

ErrHndl:
  'SQL execution errors
  'trasactions will be rolled back rollb-> true
  'current query failed to execute failed -> true
  rollB = True
  failed = True
  
  'loop through all errors and
  'save it's no., descr., source, and query in
  'the status public var
  For Each dErr In Errors
    STATUS = STATUS & "no.: " & dErr.Number & vbCrLf & _
                      "descr.: " & dErr.Description & vbCrLf & _
                      "source: " & dErr.Source & vbCrLf & _
                      "query: " & Queries(i) & vbCrLf & vbCrLf
  Next dErr
  
  'try next query
Resume Next

rBack:
  
  'at least one query failed - add the error to the status message
  STATUS = "Error" & vbCrLf & STATUS
  
  'rollback transations
  wrkSpace.Rollback
  
  'free memory, close connections
  Set dbSource = Nothing
  Set dbTarget = Nothing
  Set wrkSpace = Nothing
  
Exit Sub
  
End Sub


Sub exportLogFile()
  Dim fileOpen As String
  Dim fileNum As Integer
  Dim browseForFolderRes As Variant
  
  'after query executions - ask if user wants to save a log file
  'max msgbox prompt lenght = 1024, if status is longer troncate at 800
  'else just print status as is
  If Len(STATUS) > 800 Then
    If MsgBox(Left(STATUS, 800) & vbCrLf & "[...]" & vbCrLf & "Export log file? (recomended as full status is not shown here)", vbYesNo) = vbNo Then
      Exit Sub
    End If
  Else
    If MsgBox(STATUS & vbCrLf & vbCrLf & "Export log file?", vbYesNo) = vbNo Then
      Exit Sub
    End If
  End If
  
Retry:
  'get log export path
  browseForFolderRes = BrowseForFolder()
  
  'if path was invalied (ie user canceled),
  'ask to retry or cancel
  If browseForFolderRes = False Then
    If MsgBox("Invalid Path", vbRetryCancel) = vbRetry Then
      GoTo Retry
    Else
      Exit Sub
    End If
  End If
  
  'generate file full name:
  ': is replaced to a _ in now() to create valid path
  fileOpen = browseForFolderRes & "\excel to access log - " & Replace(Now(), ":", "_") & ".txt"
  
  fileNum = FreeFile()
  
  'open file and print out the
  'status and the executed queries
  Open fileOpen For Output As fileNum
    Print #1, "Log File - Excel to Access - " & Now()
    Print #1, "Status"
    Print #1, STATUS
    Print #1, vbCrLf, vbCrLf
    Print #1, "Executed Queries"
    If Len(EXECUTED_QUERIES) = 0 Then
      Print #1, "No queries were successfully executed"
    Else
      Print #1, EXECUTED_QUERIES
    End If
  Close #fileNum

End Sub

Private Function BrowseForFolder(Optional OpenAt As Variant) As Variant
    Dim ShellApp As Object
    
    'create shell aplication
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Please choose a folder", 0, OpenAt)
    
    'get path
    On Error Resume Next
      BrowseForFolder = ShellApp.self.path
    On Error GoTo 0
    
    'check that path is valid
    'either a local drive, with : as 2nd character
    'or a network drive, starts with \\
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


