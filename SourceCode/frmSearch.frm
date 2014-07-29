VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ファイル検索"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7335
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   3.82920e5
      TabIndex        =   4
      Top             =   10620
      Width           =   0
   End
   Begin VB.ListBox lstFiles 
      Height          =   5460
      Left            =   150
      TabIndex        =   3
      Top             =   600
      Width           =   7095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "検索"
      Height          =   375
      Left            =   6060
      TabIndex        =   1
      Top             =   120
      Width           =   1185
   End
   Begin VB.TextBox txtFile 
      Height          =   315
      Left            =   1230
      TabIndex        =   0
      Top             =   150
      Width           =   4830
   End
   Begin VB.Label Label1 
      Caption         =   "ファイル名"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   180
      Width           =   1035
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdSearch_Click()
On Error GoTo err_line
    Set gSearchedComponents = New Collection
    Me.lstFiles.Clear
    Dim i As Integer
    With gVbInstance.ActiveVBProject.VBComponents
        If txtFile <> Empty Then
            For i = 1 To .count
                If InStr(UCase(.Item(i).Name), UCase(txtFile)) <> 0 Then
                    Me.lstFiles.AddItem .Item(i).Name
                    Call gSearchedComponents.Add(.Item(i), CStr(Me.lstFiles.ListCount))
                End If
            Next
        Else
            For i = 1 To .count
                Me.lstFiles.AddItem .Item(i).Name
                Call gSearchedComponents.Add(.Item(i), CStr(Me.lstFiles.ListCount))
            Next
        End If
    End With
err_line:
    Exit Sub
End Sub

Private Sub Form_Activate()
    Call SetTopMostWindow(Me.hwnd, False)
End Sub

Private Sub Form_GotFocus()
    Me.txtFile.SetFocus
    Me.txtFile.SelStart = 0
    Me.txtFile.SelLength = Len(Me.txtFile)
End Sub

Private Sub Form_Load()
    Call SetTopMostWindow(Me.hwnd, True)
    Set gSearchedComponents = New Collection
    Dim i As Integer
    If Not gVbInstance Is Nothing Then
        If Not gVbInstance.ActiveVBProject Is Nothing Then
            With gVbInstance.ActiveVBProject.VBComponents
                For i = 1 To .count
                    Me.lstFiles.AddItem .Item(i).Name
                    Call gSearchedComponents.Add(.Item(i), CStr(Me.lstFiles.ListCount))
                Next
            End With
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Me.Hide
End Sub

Private Sub lstFiles_DblClick()
On Error GoTo err_line
    gSearchedComponents(lstFiles.ListIndex + 1).Activate
err_line:
    Me.Hide
End Sub

Private Sub lstFiles_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstFiles_DblClick
    End If
End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearch_Click
    End If
End Sub
