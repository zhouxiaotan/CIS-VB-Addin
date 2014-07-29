VERSION 5.00
Begin VB.Form frmPackage 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "成果物作成"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   7230
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox txtSource 
      Height          =   330
      Left            =   60
      TabIndex        =   13
      Top             =   420
      Width           =   2415
   End
   Begin VB.TextBox txtPackage 
      Height          =   405
      Left            =   3870
      TabIndex        =   11
      Top             =   2040
      Width           =   3225
   End
   Begin VB.DirListBox dirDest 
      Height          =   5130
      Left            =   90
      TabIndex        =   7
      Top             =   1080
      Width           =   2355
   End
   Begin VB.PictureBox picBar 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   2520
      ScaleHeight     =   105
      ScaleWidth      =   4515
      TabIndex        =   1
      Top             =   3120
      Width           =   4545
      Begin VB.Label lblBar 
         BackColor       =   &H0080FF80&
         Height          =   225
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2865
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   500
      Left            =   3570
      Top             =   4950
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   500
      Left            =   4020
      Top             =   4950
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   500
      Left            =   4470
      Top             =   4950
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   500
      Left            =   4950
      Top             =   4950
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   3150
      Top             =   4950
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "作成"
      Height          =   375
      Left            =   5850
      TabIndex        =   0
      Top             =   2670
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "オプション"
      Height          =   1755
      Left            =   2550
      TabIndex        =   3
      Top             =   150
      Width           =   4545
      Begin VB.CheckBox chkVbp 
         Caption         =   "vbpファイルを含む"
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   1290
         Value           =   1  'ﾁｪｯｸ
         Width           =   2085
      End
      Begin VB.CheckBox chksrc 
         Caption         =   "srcフォルダをルートになる(Source\src)"
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   930
         Value           =   1  'ﾁｪｯｸ
         Width           =   3555
      End
      Begin VB.CheckBox chkTest 
         Caption         =   "テストフォルダ含む(Test)"
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Value           =   1  'ﾁｪｯｸ
         Width           =   2355
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "ドキュメントフォルダを含む(Document)"
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   270
         Value           =   1  'ﾁｪｯｸ
         Width           =   4035
      End
   End
   Begin VB.TextBox lstFiles 
      Height          =   2865
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   3  '両方
      TabIndex        =   15
      Top             =   3330
      Width           =   4545
   End
   Begin VB.Label Label3 
      Caption         =   "保存元ルートフォルダ"
      Height          =   225
      Left            =   90
      TabIndex        =   14
      Top             =   180
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "保存先"
      Height          =   225
      Left            =   90
      TabIndex        =   12
      Top             =   810
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "パッケージID"
      Height          =   255
      Left            =   2580
      TabIndex        =   10
      Top             =   2070
      Width           =   1275
   End
   Begin VB.Label lblMsg 
      BackColor       =   &H80000008&
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2550
      TabIndex        =   9
      Top             =   2700
      Width           =   3075
   End
End
Attribute VB_Name = "frmPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LB_SETTABSTOPS As Long = &H192&
Private IsSearching As Boolean
Private FoundCount As Long
Private SearchedFile As Long
Private CompletedTimer As Integer
Private StrDest As String
Private MaxTimerCount As Integer

Private myFSO As Scripting.FileSystemObject

Private Sub cmdSearch_Click()

    
    StrDest = dirDest.Path & "\" & txtPackage
    If Not myFSO.FolderExists(StrDest) Then
        Call VBA.MkDir(StrDest)
    End If
    
    If chkDoc.Value And Not myFSO.FolderExists(StrDest & "\Document") Then
        Call VBA.MkDir(StrDest & "\Document")
    End If
    
    If chkTest.Value And Not myFSO.FolderExists(StrDest & "\Test") Then
        Call VBA.MkDir(StrDest & "\Test")
    End If
    
    StrDest = StrDest & "\Source\"
    If Not myFSO.FolderExists(StrDest) Then
        Call VBA.MkDir(StrDest)
    End If

    
    If chksrc.Value Then
        StrDest = StrDest & "\src\"
        If Not myFSO.FolderExists(StrDest) Then
            Call VBA.MkDir(StrDest)
        End If
    End If
    
    gDestDir = dirDest.Path
    gSorcDir = txtSource
    Call WriteINI(INI_FILE, SECTION, DESTDIR, gDestDir)
    Call WriteINI(INI_FILE, SECTION, SORCDIR, gSorcDir)
    
    Dim i As Integer, j As Integer
    Dim rootPath  As String, rootPath2()  As String
    rootPath = Replace(gVbInstance.ActiveVBProject.filename, gSorcDir, StrDest)
    Dim strTemp As String
    If Not myFSO.FolderExists(rootPath) Then
        rootPath2 = Split(rootPath, "\")
        strTemp = Empty
        For j = LBound(rootPath2) To UBound(rootPath2) - 1
            strTemp = strTemp & IIf(j = LBound(rootPath2), "", "\") & rootPath2(j)
            If Not myFSO.FolderExists(strTemp) Then
                Call myFSO.CreateFolder(strTemp)
            End If
        Next
    End If
    FoundCount = 0
    If IsSearching = False Then
        Set gSearchedTexts = New Collection
        Me.lstFiles.Text = Empty
        lblBar.width = 0
        SearchedFile = 0
        CompletedTimer = 0
        If chkVbp.Value Then
            Call myFSO.CopyFile(gVbInstance.ActiveVBProject.filename, rootPath)
            Call AddText(gVbInstance.ActiveVBProject.filename)
        End If
        MaxTimerCount = 4
        IsSearching = True
        For i = 0 To MaxTimerCount
            Timer1(i).Enabled = True
            DoEvents
        Next
        cmdSearch.Caption = "終止"
        lblMsg.Caption = "作成中..."
    Else
        cmdSearch.Caption = "作成"
        IsSearching = False
        lblMsg.Caption = "終止しました。"
    End If
End Sub


Private Sub Form_Activate()
    Call SetTopMostWindow(Me.hwnd, False)
End Sub

Private Sub Form_DblClick()
    Me.Hide
End Sub

Private Sub Form_Load()
    Call SetTopMostWindow(Me.hwnd, True)
    Set gSearchedTexts = New Collection
   
    lblBar.width = 0
    SearchedFile = 0
    
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    
    If Not myFSO.FolderExists(gDestDir) Then
        gDestDir = "C:\"
        Call WriteINI(INI_FILE, SECTION, DESTDIR, gDestDir)
    End If
    
    dirDest.Path = gDestDir
    txtSource = gSorcDir
    

    txtPackage = gPackageId
    
End Sub

Private Sub Timer1_Timer(Index As Integer)
On Err GoTo error
   Timer1(Index).Enabled = False
   With gVbInstance.ActiveVBProject.VBComponents
        Dim startIdx As Integer, endIdx As Integer
        Dim objLines() As String
        Dim rootPath  As String, rootPath2()  As String
        Dim strFolder As String
        Dim filename As String, filename2 As String, frxname As String, vbppath As String
        
        startIdx = (.count \ (MaxTimerCount + 1)) * Index
        endIdx = (.count \ (MaxTimerCount + 1)) * (Index + 1)
        If endIdx > .count Then endIdx = .count
        If Index = MaxTimerCount Then endIdx = .count
        If txtPackage <> Empty Then
            For i = startIdx + 1 To endIdx
                If .Item(i).Type <> vbext_ComponentType.vbext_ct_RelatedDocument Then
                    If Not .Item(i).CodeModule Is Nothing Then
                        If .Item(i).CodeModule.Find(txtPackage, 1, 1, -1, -1, False, False) Then
                           rootPath = Replace(.Item(i).FileNames(1), gSorcDir, StrDest)
                           objLines = Split(rootPath, "\")
                           strFolder = Empty
                           For j = LBound(objLines) To UBound(objLines) - 1
                                strFolder = strFolder & IIf(j = LBound(objLines), "", "\") & objLines(j)
                                If Not myFSO.FolderExists(strFolder) Then
                                    Call myFSO.CreateFolder(strFolder)
                                End If
                           Next
                           Call myFSO.CopyFile(.Item(i).FileNames(1), rootPath)
                           
                           If InStr(1, objLines(UBound(objLines)), ".frm", vbTextCompare) > 0 Then
                                filename2 = Replace(rootPath, ".frm", ".frx")
                                frxname = Replace(.Item(i).FileNames(1), ".frm", ".frx")
                                If myFSO.FileExists(frxname) Then
                                    Call myFSO.CopyFile(frxname, filename2)
                                End If
                           End If
                           Call AddText(.Item(i).FileNames(1))
                        End If
                    End If
                End If
                SearchedFile = SearchedFile + 1
                lblBar.width = SearchedFile / .count * picBar.width
                If IsSearching = False Then Exit For
            Next
        End If
        CompletedTimer = CompletedTimer + 1
        If CompletedTimer > MaxTimerCount Then
            cmdSearch.Caption = "作成"
            IsSearching = False
            lblMsg.Caption = "作成完了。"
        End If
        Debug.Print Index & vbTab & .count & vbTab & startIdx & vbTab & endIdx
    End With
    Debug.Print "TimerIndex: " & Index
error:
    Exit Sub
End Sub

Private Sub AddText(strFileName As String)
    lstFiles.Text = lstFiles.Text & Replace(strFileName, txtSource, "") & vbCrLf
End Sub
