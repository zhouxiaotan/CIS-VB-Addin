VERSION 5.00
Begin VB.Form frmAddIn 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'ŒÅ’èÀÞ²±Û¸Þ
   Caption         =   "CIS Add-In for Visual Basic Joel 1.o"
   ClientHeight    =   3285
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '‰æ–Ê‚Ì’†‰›
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H80000012&
      Height          =   180
      Left            =   2340
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   1410
      Width           =   705
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Height          =   555
      Left            =   1530
      TabIndex        =   6
      Top             =   1170
      Width           =   2295
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Height          =   1425
      Left            =   1020
      TabIndex        =   3
      Top             =   750
      Width           =   3285
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Height          =   945
         Left            =   210
         TabIndex        =   4
         Top             =   240
         Width           =   2865
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Height          =   1815
      Left            =   750
      TabIndex        =   2
      Top             =   570
      Width           =   3855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Height          =   2325
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   2865
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   5025
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private i As Integer

Private Sub cmdDone_Click()
    Call SetTopMostWindow(Me.hwnd, False)
    Me.Hide
End Sub

Private Sub Form_Load()
    Call SetTopMostWindow(Me.hwnd, True)
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label3.Alignment = i Mod 3
    If i > 500 Then i = 0
    i = i + 1
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    canel = 1
    Me.Hide
End Sub

Private Sub Frame1_DblClick()
    Me.Hide
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = 4
End Sub
