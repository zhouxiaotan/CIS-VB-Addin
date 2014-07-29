VERSION 5.00
Begin VB.Form frmGoto 
   BorderStyle     =   4  'å≈íË¬∞Ÿ ≥®›ƒﬁ≥
   Caption         =   "Goto"
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1860
      TabIndex        =   2
      Top             =   150
      Width           =   855
   End
   Begin VB.TextBox txtRow 
      Alignment       =   1  'âEëµÇ¶
      Height          =   330
      Left            =   930
      TabIndex        =   0
      Text            =   "0"
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "çsî‘çÜÅF"
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err_line
    Me.Hide
    If IsNumeric(txtRow) Then
        If Not gVbInstance.ActiveCodePane Is Nothing Then
            With gVbInstance.ActiveCodePane
                Dim line As Integer
                line = CInt(txtRow)
                If line <= 0 Then line = 1
                Call .CodeModule.CodePane.SetSelection(line, 1, line, 1)
            End With
        End If
    End If
err_line:
    Exit Sub
End Sub

Private Sub Form_Activate()
    Call SetTopMostWindow(Me.hwnd, False)
End Sub

Private Sub Form_DblClick()
    Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Me.Hide
End Sub

Private Sub txtRow_GotFocus()
    Me.txtRow.SelLength = Len(Me.txtRow)
End Sub

Private Sub txtRow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Command1_Click
    End If
End Sub

Private Sub Form_Load()
    Call SetTopMostWindow(Me.hwnd, True)
End Sub
