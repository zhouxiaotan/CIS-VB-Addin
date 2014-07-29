VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   11460
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   18105
   _ExtentX        =   31935
   _ExtentY        =   20214
   _Version        =   393216
   Description     =   "CodeLineCounter Add-In"
   DisplayName     =   "CodeLineCounter"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public VbInstance            As VBIDE.VBE

Public WithEvents SetHandler As CommandBarEvents
Attribute SetHandler.VB_VarHelpID = -1
Public WithEvents AddHandler As CommandBarEvents
Attribute AddHandler.VB_VarHelpID = -1
Public WithEvents DelHandler As CommandBarEvents
Attribute DelHandler.VB_VarHelpID = -1
Public WithEvents RepHandler As CommandBarEvents
Attribute RepHandler.VB_VarHelpID = -1
Public WithEvents TilHandler As CommandBarEvents
Attribute TilHandler.VB_VarHelpID = -1
Public WithEvents AbtHandler As CommandBarEvents
Attribute AbtHandler.VB_VarHelpID = -1
Public WithEvents SrhHandler As CommandBarEvents
Attribute SrhHandler.VB_VarHelpID = -1
Public WithEvents TxtHandler As CommandBarEvents
Attribute TxtHandler.VB_VarHelpID = -1
Public WithEvents PkgHandler As CommandBarEvents
Attribute PkgHandler.VB_VarHelpID = -1
Public WithEvents ClsHandler As CommandBarEvents
Attribute ClsHandler.VB_VarHelpID = -1
Public WithEvents RdoHandler As CommandBarEvents
Attribute RdoHandler.VB_VarHelpID = -1
Public WithEvents OpdHandler As CommandBarEvents
Attribute OpdHandler.VB_VarHelpID = -1
Public WithEvents GotHandler As CommandBarEvents
Attribute GotHandler.VB_VarHelpID = -1

Private mMidCmd As Office.CommandBarButton

Private lineLeft As Long
Private lineRight As Long
Private lineTop As Long
Private lineBottom As Long

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    Set VbInstance = Application
    Set gVbInstance = VbInstance
    
    Call ReadAll
    
    Call AddToAddInCommandBar
    
    Exit Sub
    
error_handler:
    
    MsgBox Err.DESCRIPTION
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    Unload frmAddIn
    Unload frmSetting
    Unload frmSearch
    Unload frmSearchText
    Unload frmPackage
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)

End Sub

Private Sub AddToAddInCommandBar()
    Dim barCis As CommandBar
    
    Dim cmdSetting As Office.CommandBarButton
    Dim cmdAdd As Office.CommandBarButton
    Dim cmdReplace As Office.CommandBarButton
    Dim cmdDelete As Office.CommandBarButton
    Dim cmdTitle As Office.CommandBarButton
    Dim cmdAbout As Office.CommandBarButton
    Dim cmdSearch As Office.CommandBarButton
   
    Dim cmdPackage As Office.CommandBarButton
    Dim cmdCloseAll As Office.CommandBarButton
    Dim cmdRedo As Office.CommandBarButton
    Dim cmdOpenFolder As Office.CommandBarButton
    Dim cmdGotoLine As Office.CommandBarButton
   
On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set barCis = VbInstance.CommandBars.Add("CIS", msoBarTop, False, True)
    If barCis Is Nothing Then
        MsgBox "Can't Create CommandBar"
        Exit Sub
    End If
    
    barCis.Visible = True
    
    'Add it to the command bar
    Set cmdSetting = barCis.Controls.Add(msoControlButton, 601, "Setting")
    Set cmdAdd = barCis.Controls.Add(msoControlButton, 80, "Add")
    Set cmdReplace = barCis.Controls.Add(msoControlButton, 97, "Replace")
    Set cmdDelete = barCis.Controls.Add(msoControlButton, 83, "Delete")
    Set cmdTitle = barCis.Controls.Add(msoControlButton, 99, "Title")
    Set cmdRedo = barCis.Controls.Add(msoControlButton, 37, "Redo")
    Set cmdSearch = barCis.Controls.Add(msoControlButton, 183, "Search")
    Set cmdSearchText = barCis.Controls.Add(msoControlButton, 141, "Search")
    Set cmdPackage = barCis.Controls.Add(msoControlButton, 577, "Package")
    Set cmdCloseAll = barCis.Controls.Add(msoControlButton, 52, "Close All")
    Set cmdOpenFolder = barCis.Controls.Add(msoControlButton, 721, "Open Folder")
    Set cmdGotoLine = barCis.Controls.Add(msoControlButton, 44, "Go To")
    Set cmdAbout = barCis.Controls.Add(msoControlButton, 352, "About")
    
    cmdSetting.Caption = "Setting"
    cmdAdd.Caption = "Appending"
    cmdReplace.Caption = "Changing"
    cmdDelete.Caption = "Deleting"
    cmdTitle.Caption = "Title"
    cmdAbout.Caption = "About..."
    cmdSearch.Caption = "File Searching"
    cmdSearchText.Caption = "Text Searching"
    cmdPackage.Caption = "Create Package"
    cmdCloseAll.Caption = "Close All"
    cmdRedo.Caption = "Redo Package"
    cmdOpenFolder.Caption = "Open Folder For Selected File"
    cmdGotoLine.Caption = "GoTo Line"
    
    Set Me.SetHandler = VbInstance.Events.CommandBarEvents(cmdSetting)
    Set Me.AddHandler = VbInstance.Events.CommandBarEvents(cmdAdd)
    Set Me.RepHandler = VbInstance.Events.CommandBarEvents(cmdReplace)
    Set Me.DelHandler = VbInstance.Events.CommandBarEvents(cmdDelete)
    Set Me.TilHandler = VbInstance.Events.CommandBarEvents(cmdTitle)
    Set Me.AbtHandler = VbInstance.Events.CommandBarEvents(cmdAbout)
    Set Me.SrhHandler = VbInstance.Events.CommandBarEvents(cmdSearch)
    Set Me.TxtHandler = VbInstance.Events.CommandBarEvents(cmdSearchText)
    Set Me.PkgHandler = VbInstance.Events.CommandBarEvents(cmdPackage)
    Set Me.ClsHandler = VbInstance.Events.CommandBarEvents(cmdCloseAll)
    Set RdoHandler = VbInstance.Events.CommandBarEvents(cmdRedo)
    Set OpdHandler = VbInstance.Events.CommandBarEvents(cmdOpenFolder)
    Set GotHandler = VbInstance.Events.CommandBarEvents(cmdGotoLine)
    
    Exit Sub
    
AddToAddInCommandBarErr:
    MsgBox Err.DESCRIPTION
End Sub

Private Sub Operation(nFlag As Integer)
    If Not VbInstance.ActiveVBProject.VBE.ActiveCodePane Is Nothing Then
        Call VbInstance.ActiveVBProject.VBE.ActiveCodePane.GetSelection(lineTop, lineLeft, lineBottom, lineRight)
        Dim i As Integer
        Dim lineCount As Integer
        Dim line As String
        With VbInstance.ActiveVBProject.VBE.ActiveCodePane.CodeModule
            If nFlag = 1 Or nFlag = 2 Then
                If lineTop <> lineBottom And lineRight = 1 Then
                    lineBottom = lineBottom - 1
                End If
            End If
            lineCount = lineBottom - lineTop
            For i = 0 To lineCount
                line = .Lines(lineTop + i, 1)
                If i = lineCount Then
                    If line = Empty Then
                        Exit For
                    Else
                        lineBottom = lineBottom + 1
                    End If
                End If
                If nFlag <> 3 Then
                    Call .ReplaceLine(lineTop + i, "''" & line)
                End If
            Next
            If nFlag = 1 Then
                Call .InsertLines(lineBottom, DEL_STRING_E)
                Call .InsertLines(lineTop, DEL_STRING_S)
            ElseIf nFlag = 2 Then
                
                Call .InsertLines(lineBottom, REP_STRING_E)
                Call .InsertLines(lineTop, REP_STRING_S)
                Call .CodePane.SetSelection(lineBottom + 2, 1, lineBottom + 2, 1)
            ElseIf nFlag = 3 Then
                Call .InsertLines(lineBottom, ADD_STRING_E)
                Call .InsertLines(lineTop, ADD_STRING_S)
            End If
        End With
  End If
End Sub

Private Sub ClsHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo err_line
    Set mMidCmd = CommandBarControl
    mMidCmd.Enabled = False
    DoEvents
    Dim obj As Window
    For Each obj In gVbInstance.ActiveVBProject.VBE.Windows
        If (obj.Type = vbext_wt_CodeWindow Or obj.Type = vbext_wt_Designer) And obj.Visible Then
            obj.Close
        End If
    Next
err_line:
    mMidCmd.Enabled = True
    Exit Sub
End Sub

Private Sub DelHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo err_line
    Set mMidCmd = CommandBarControl
    mMidCmd.Enabled = False
    DoEvents
    Call Operation(1)
err_line:
    mMidCmd.Enabled = True
    Exit Sub
End Sub

Private Sub GotHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo err_line
    Set mMidCmd = CommandBarControl
    mMidCmd.Enabled = False
    DoEvents
    frmGoto.Show
    frmGoto.txtRow.SetFocus
    frmGoto.txtRow.SelStart = 0
    frmGoto.txtRow.SelLength = Len(frmGoto.txtRow)
err_line:
    mMidCmd.Enabled = True
    Exit Sub
End Sub

Private Sub OpdHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo err_line
    Set mMidCmd = CommandBarControl
    mMidCmd.Enabled = False
    DoEvents
    Dim i As Integer
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    If Not VbInstance.ActiveVBProject.VBE.SelectedVBComponent Is Nothing Then
        With VbInstance.ActiveVBProject.VBE.SelectedVBComponent
             Call Shell("explorer.exe /select,""" & .FileNames(1) & """", vbNormalFocus)
        End With
    End If
err_line:
    mMidCmd.Enabled = True
    Exit Sub
End Sub

Private Sub PkgHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo err_line
    Set mMidCmd = CommandBarControl
    mMidCmd.Enabled = False
    DoEvents
    Call SetParent(frmPackage.hwnd, VbInstance.MainWindow.hwnd)
    Call frmPackage.Show
err_line:
    mMidCmd.Enabled = True
    Exit Sub
End Sub

Private Sub RdoHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo err_line
    Dim i As Integer
    Dim blnRemoving As Boolean
    Set mMidCmd = CommandBarControl
    mMidCmd.Enabled = False
    DoEvents
    If Not VbInstance.ActiveVBProject.VBE.ActiveCodePane Is Nothing Then
        With VbInstance.ActiveVBProject.VBE.ActiveCodePane
        Call .GetSelection(lineTop, lineLeft, lineBottom, lineRight)
            For i = lineTop To lineBottom
                If InStr(1, .CodeModule.Lines(i, 1), "#IF", vbTextCompare) Then
                    If Right(Trim(.CodeModule.Lines(i, 1)), 1) = "D" Then
                        GoTo Delete
                    ElseIf Right(Trim(.CodeModule.Lines(i, 1)), 1) = "R" Then
                        GoTo Replace
                    End If
                ElseIf Right(Trim(.CodeModule.Lines(i, 1)), 1) = "A" And Left(Trim(.CodeModule.Lines(i, 1)), 2) = "'@" Then
                    GoTo Add
                End If
            Next
        
Replace:
            For i = lineBottom To lineTop Step -1
                If Left(.CodeModule.Lines(i, 1), 3) = "#If" And Right(Trim(.CodeModule.Lines(i, 1)), 1) = "R" Then
                    Call .CodeModule.DeleteLines(i)
                ElseIf Left(.CodeModule.Lines(i, 1), 4) = "#End" And Right(Trim(.CodeModule.Lines(i, 1)), 1) = "E" Then
                    Call .CodeModule.DeleteLines(i)
                    blnRemoving = True
                ElseIf Left(.CodeModule.Lines(i, 1), 5) = "#Else" Then
                    Call .CodeModule.DeleteLines(i)
                    blnRemoving = False
                ElseIf (Not blnRemoving) And Left(.CodeModule.Lines(i, 1), 1) = "'" Then
                    Call .CodeModule.ReplaceLine(i, Mid(.CodeModule.Lines(i, 1), 3))
                ElseIf blnRemoving Then
                    Call .CodeModule.DeleteLines(i)
                End If
            Next
            GoTo ExitSub
Delete:
            For i = lineBottom To lineTop Step -1
                If Left(.CodeModule.Lines(i, 1), 4) = "#End" And Right(Trim(.CodeModule.Lines(i, 1)), 1) = "E" Then
                    Call .CodeModule.DeleteLines(i)
                    blnRemoving = True
                ElseIf Left(.CodeModule.Lines(i, 1), 3) = "#If" And Right(Trim(.CodeModule.Lines(i, 1)), 1) = "D" Then
                    Call .CodeModule.DeleteLines(i)
                    blnRemoving = False
                ElseIf blnRemoving Then
                    Call .CodeModule.ReplaceLine(i, Mid(.CodeModule.Lines(i, 1), 3))
                End If
            Next
            GoTo ExitSub
Add:
            For i = lineBottom To lineTop Step -1
                If Left(.CodeModule.Lines(i, 1), 2) = "'@" And Right(Trim(.CodeModule.Lines(i, 1)), 1) = "E" Then
                    Call .CodeModule.DeleteLines(i)
                    blnRemoving = True
                ElseIf Left(.CodeModule.Lines(i, 1), 1) = "'@" And Right(Trim(.CodeModule.Lines(i, 1)), 1) = "A" Then
                    Call .CodeModule.DeleteLines(i)
                    blnRemoving = False
                ElseIf blnRemoving Then
                    Call .CodeModule.DeleteLines(i)
                End If
            Next
        End With
    End If
ExitSub:
err_line:
    mMidCmd.Enabled = True
    Exit Sub
End Sub

Private Sub RepHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo err_line
    Set mMidCmd = CommandBarControl
    mMidCmd.Enabled = False
    DoEvents
    Call Operation(2)
err_line:
    mMidCmd.Enabled = True
    Exit Sub
End Sub

Private Sub SetHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo err_line
    Set mMidCmd = CommandBarControl
    mMidCmd.Enabled = False
    DoEvents
    Call SetParent(frmSetting.hwnd, VbInstance.MainWindow.hwnd)
    Call frmSetting.Show
err_line:
    mMidCmd.Enabled = True
    Exit Sub
End Sub

Private Sub SrhHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo err_line
    Set mMidCmd = CommandBarControl
    mMidCmd.Enabled = False
    DoEvents
    Call SetParent(frmSearch.hwnd, VbInstance.MainWindow.hwnd)
    frmSearch.Show
err_line:
    mMidCmd.Enabled = True
    Exit Sub
End Sub

Private Sub TilHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo err_line
    Set mMidCmd = CommandBarControl
    mMidCmd.Enabled = False
    DoEvents
    If Not VbInstance.ActiveVBProject.VBE.ActiveCodePane Is Nothing Then
        Call VbInstance.ActiveVBProject.VBE.ActiveCodePane.GetSelection(lineTop, lineLeft, lineBottom, lineRight)
        Call VbInstance.ActiveVBProject.VBE.ActiveCodePane.CodeModule.InsertLines(lineTop, TIL_STRING)
    End If
err_line:
    mMidCmd.Enabled = True
    Exit Sub
End Sub

Private Sub AddHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo err_line
    Set mMidCmd = CommandBarControl
    mMidCmd.Enabled = False
    DoEvents
    If Not VbInstance.ActiveVBProject.VBE.ActiveCodePane Is Nothing Then
        Call VbInstance.ActiveVBProject.VBE.ActiveCodePane.GetSelection(lineTop, lineLeft, lineBottom, lineRight)
        With VbInstance.ActiveVBProject.VBE.ActiveCodePane
            If lineTop = lineBottom And lineLeft = lineRight Then
                If lineTop < .CodeModule.CountOfLines Then lineTop = lineTop + 1
                Call .CodeModule.InsertLines(lineTop, ADD_STRING_E)
                Call .CodeModule.InsertLines(lineTop, "")
                Call .CodeModule.InsertLines(lineTop, ADD_STRING_S)
                Call .CodeModule.CodePane.SetSelection(lineTop + 2, 1, lineTop + 2, 1)
            Else
                Call Operation(3)
            End If
        End With
    End If
err_line:
    mMidCmd.Enabled = True
    Exit Sub
End Sub

Private Sub AbtHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
On Error GoTo err_line
    Set mMidCmd = CommandBarControl
    mMidCmd.Enabled = False
    DoEvents
    Call SetParent(frmAddIn.hwnd, VbInstance.MainWindow.hwnd)
    Call frmAddIn.Show
err_line:
    mMidCmd.Enabled = True
    Exit Sub
End Sub

Private Sub TxtHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim strSelectedText As String
    Dim colstart As Long, colend As Long, rowstart As Long, rowend As Long
On Error GoTo err_line
    Set mMidCmd = CommandBarControl
    mMidCmd.Enabled = False
    DoEvents
    Call SetParent(frmSearchText.hwnd, VbInstance.MainWindow.hwnd)
    Call gVbInstance.ActiveCodePane.GetSelection(rowstart, colstart, rowend, colend)
    strSelectedText = gVbInstance.ActiveCodePane.CodeModule.Lines(rowstart, 1)
    colstart = convertColumn2Pos(strSelectedText, colstart)
    colend = convertColumn2Pos(strSelectedText, colend)
    frmSearchText.txtFile.Text = Mid(strSelectedText, colstart, colend - colstart)
    Call frmSearchText.Show
err_line:
    mMidCmd.Enabled = True
    Exit Sub
End Sub
