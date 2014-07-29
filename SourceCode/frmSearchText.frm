VERSION 5.00
Begin VB.Form frmSearchText 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "テキスト検索"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   13485
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdClose 
      Caption         =   "閉じる"
      Height          =   405
      Left            =   12390
      TabIndex        =   6
      Top             =   5820
      Width           =   1065
   End
   Begin VB.PictureBox picBar 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      ScaleHeight     =   225
      ScaleWidth      =   13245
      TabIndex        =   4
      Top             =   540
      Width           =   13275
      Begin VB.Label lblBar 
         BackColor       =   &H0080FF80&
         Height          =   225
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   7335
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   500
      Left            =   990
      Top             =   5400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   500
      Left            =   1440
      Top             =   5400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   500
      Left            =   1890
      Top             =   5400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   500
      Left            =   2340
      Top             =   5400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   570
      Top             =   5400
   End
   Begin VB.ListBox lstFiles 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4890
      Left            =   150
      TabIndex        =   3
      Top             =   870
      Width           =   13245
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "検索"
      Height          =   375
      Left            =   7590
      TabIndex        =   1
      Top             =   90
      Width           =   1065
   End
   Begin VB.TextBox txtFile 
      Height          =   315
      Left            =   1380
      TabIndex        =   0
      Top             =   150
      Width           =   5985
   End
   Begin VB.Label lblAlert 
      Caption         =   "友達、とても多く、無理だ。"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5040
      TabIndex        =   9
      Top             =   5850
      Visible         =   0   'False
      Width           =   5445
   End
   Begin VB.Label lblKekka 
      Height          =   255
      Left            =   9720
      TabIndex        =   8
      Top             =   150
      Width           =   3675
   End
   Begin VB.Label Label2 
      Caption         =   "検索結果："
      Height          =   255
      Left            =   8850
      TabIndex        =   7
      Top             =   150
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "テキスト"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   180
      Width           =   1035
   End
End
Attribute VB_Name = "frmSearchText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IsSearching As Boolean
Private FoundCount As Long
Private SearchedFile As Long
Private CompletedTimer As Integer
Private MaxTimerCount As Integer
Private FoundFileCount As Long

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdSearch_Click()
On Error GoTo err_line
     Dim i As Integer, j As Integer
     Dim count As Integer
     FoundCount = 0
     If IsSearching = False Then
        Set gSearchedTexts = New Collection
        Me.lstFiles.Clear
        lblBar.width = 0
        SearchedFile = 0
        FoundFileCount = 0
        CompletedTimer = 0
        lblKekka.Caption = Empty
        MaxTimerCount = 4
        lblAlert.Visible = False
        For i = 0 To MaxTimerCount
            Timer1(i).Enabled = True
        Next
        cmdSearch.Caption = "終止"
        IsSearching = True
    Else
        cmdSearch.Caption = "検索"
        IsSearching = False
    End If
    Exit Sub
err_line:
    cmdSearch.Caption = "検索"
    IsSearching = False
End Sub

Private Sub Form_Activate()
    Call SetTopMostWindow(Me.hwnd, False)
    Me.txtFile.SetFocus
    lblAlert.Visible = False
End Sub

Private Sub Form_Load()
    Call SetTopMostWindow(Me.hwnd, True)
    Set gSearchedTexts = New Collection
    Dim f1 As Integer, f2 As Integer
    f1 = 100
    f2 = 130
    
    ReDim TabStops(1) As Long
    TabStops(0) = f1
    TabStops(1) = f2
    
    Call SendMessage(lstFiles.hwnd, LB_SETTABSTOPS, 2&, TabStops(0))
    
    lblBar.width = 0
    SearchedFile = 0

    Call DisplayHScroll(lstFiles)
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.Hide
    Cancel = 1
End Sub

Private Sub lstFiles_DblClick()
On Error GoTo err_line
    Dim com As VBComponent
    Dim lngLine As Long
    Dim strText As String
    Dim intStart As Integer
    
    Set com = gSearchedTexts(lstFiles.ListIndex + 1)
    com.CodeModule.CodePane.Show
    lngLine = Split(lstFiles.Text, vbTab)(1)
    strText = com.CodeModule.Lines(lngLine, 1)
    intStart = InStrB(1, StrConv(strText, vbFromUnicode), StrConv(txtFile, vbFromUnicode), vbTextCompare)
    If intStart <> 0 Then
        Call com.CodeModule.CodePane.SetSelection(lngLine, intStart, lngLine, intStart + LenB(StrConv(txtFile, vbFromUnicode)))
    Else
        Call com.CodeModule.CodePane.SetSelection(lngLine, 1, lngLine, LenB(StrConv(strText, vbFromUnicode)) + 1)
    End If
err_line:
    Me.Hide
End Sub

Private Sub lstFiles_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstFiles_DblClick
    End If
End Sub

Private Sub Timer1_Timer(Index As Integer)
On Err GoTo error
   Timer1(Index).Enabled = False
   With gVbInstance.ActiveVBProject.VBComponents
        Dim startIdx As Integer, endIdx As Integer
        Dim objLines() As String
        startIdx = (.count \ (MaxTimerCount + 1)) * Index
        endIdx = (.count \ (MaxTimerCount + 1)) * (Index + 1)
        If endIdx > .count Then endIdx = .count
        If Index = MaxTimerCount Then endIdx = .count
        If txtFile <> Empty And startindex >= 0 Then
            For i = startIdx + 1 To endIdx
                If .Item(i).Type <> vbext_ComponentType.vbext_ct_RelatedDocument Then
                    If Not .Item(i).CodeModule Is Nothing Then
                        If .Item(i).CodeModule.Find(txtFile, 1, 1, -1, -1, False, False) Then
                            FoundFileCount = FoundFileCount + 1
                            objLines = Split(.Item(i).CodeModule.Lines(1, .Item(i).CodeModule.CountOfLines), vbCrLf)
                            For j = LBound(objLines) To UBound(objLines)
                                If Me.lstFiles.ListCount >= 32766 Then
                                    lblAlert.Visible = True
                                    cmdSearch.Caption = "検索"
                                    IsSearching = False
                                    Exit For
                                End If
                                If Not IsSearching Then
                                    cmdSearch.Caption = "検索"
                                    Exit For
                                End If
                                If InStr(1, objLines(j), txtFile, vbTextCompare) > 0 Then
                                    FoundCount = FoundCount + 1
                                    lblKekka.Caption = .count & "の" & FoundFileCount & "ファイル中に" & FoundCount & "件があった。"
                                    Me.lstFiles.AddItem .Item(i).Name & vbTab & j + 1 & vbTab & objLines(j)
                                    Call gSearchedTexts.Add(.Item(i), CStr(Me.lstFiles.ListCount))
                                    DoEvents
                                End If
                            Next
                        End If
                    End If
                End If
                SearchedFile = SearchedFile + 1
                lblBar.width = SearchedFile / .count * picBar.width
                If Not IsSearching Then Exit Sub
                If SearchedFile = .count Then
                    cmdSearch.Caption = "検索"
                    IsSearching = False
                End If
            Next
        End If
        CompletedTimer = CompletedTimer + 1
        Debug.Print Index & vbTab & .count & vbTab & startIdx & vbTab & endIdx
    End With
error:
    Exit Sub
End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearch_Click
    End If
End Sub
