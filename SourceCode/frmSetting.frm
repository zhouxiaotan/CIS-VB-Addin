VERSION 5.00
Begin VB.Form frmSetting 
   Caption         =   "設定"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "設定"
   ScaleHeight     =   4695
   ScaleWidth      =   10230
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.TextBox txtAuthor 
      Height          =   375
      Left            =   4770
      TabIndex        =   11
      Top             =   180
      Width           =   1305
   End
   Begin VB.TextBox txtOther 
      Height          =   1455
      Left            =   270
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   2490
      Width           =   9795
   End
   Begin VB.TextBox txtComment 
      Height          =   705
      Left            =   270
      TabIndex        =   8
      Top             =   1170
      Width           =   9765
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "キャンセル"
      Height          =   375
      Left            =   8730
      TabIndex        =   6
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   "TD"
      Height          =   375
      Left            =   9660
      TabIndex        =   5
      Top             =   180
      Width           =   375
   End
   Begin VB.TextBox txtDate 
      Height          =   375
      Left            =   7380
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   180
      Width           =   2235
   End
   Begin VB.TextBox txtPackageId 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   180
      Width           =   2055
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   7260
      TabIndex        =   0
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "作成者"
      Height          =   255
      Index           =   4
      Left            =   3750
      TabIndex        =   12
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "追加パッケージID"
      Height          =   255
      Index           =   3
      Left            =   270
      TabIndex        =   9
      Top             =   1980
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "説明"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   810
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "日付"
      Height          =   255
      Index           =   1
      Left            =   6300
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "パッケージID"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Call SetTopMostWindow(Me.hwnd, False)
    Me.Hide
End Sub

Private Sub cmdDate_Click()
    txtDate = Format(Now, "yyyy/MM/dd")
End Sub

Private Sub cmdOk_Click()
    modFormat.gAuthor = txtAuthor.Text
    modFormat.gPackageId = txtPackageId.Text
    modFormat.gDescription = txtComment.Text
    modFormat.gDateTime = txtDate.Text
    modFormat.gOther = txtOther.Text
    Call SetTopMostWindow(Me.hwnd, False)
    Call WriteAll
    Me.Hide
End Sub

Private Sub Form_Activate()
    Call SetTopMostWindow(Me.hwnd, False)
End Sub

Private Sub Form_Load()

    Call SetTopMostWindow(Me.hwnd, True)
    
    txtAuthor.Text = modFormat.gAuthor
    txtPackageId.Text = modFormat.gPackageId
    txtComment.Text = modFormat.gDescription
    txtDate.Text = modFormat.gDateTime
    txtOther.Text = modFormat.gOther
    
End Sub
