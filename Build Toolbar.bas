Attribute VB_Name = "modOpenClose"
Option Explicit

Public Const gsTOOLBARNAME as String = "ToolbarName"

Public Sub BuildBar()
    
    Dim cbrBar As CommandBar
    Dim ctlButton As CommandBarButton
    Dim ctlDropDown As CommandBarPopup
    
    On Error Resume Next
        Application.CommandBars(gsTOOLBARNAME).Delete
    On Error GoTo 0
    
    ' Create the command bar.
    Set cbrBar = Application.CommandBars.Add(gsTOOLBARNAME, _
                                        msoBarTop, False, True)
    cbrBar.Visible = True
    
    ' Add the controls required by our application.
    Set ctlButton = cbrBar.Controls.Add(msoControlButton)
    With ctlButton
        .Style = msoButtonIconAndCaption
        .Caption = "CAPTION"
        .FaceId = 107
        .OnAction = "module.Sub"
    End With

End Sub

Public Sub DeleteBar()
    On Error Resume Next
        Application.CommandBars(gsTOOLBARNAME).Delete
    On Error GoTo 0
End Sub


