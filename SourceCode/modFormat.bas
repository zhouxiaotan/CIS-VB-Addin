Attribute VB_Name = "modFormat"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Const GWL_STYLE = (-16)
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_FRAMECHANGED = &H20
Public Const WS_HSCROLL = &H100000
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_GETHORIZONTALEXTENT = &H193

Public Const LB_SETTABSTOPS As Long = &H192&

Public Const ADD_FORMAT_S As String = "'@#PACKAGEID#A"
Public Const ADD_FORMAT_E As String = "'@#PACKAGEID#E"
Public Const DEL_FORMAT_S As String = "#If 0 Then '@#PACKAGEID#D"
Public Const DEL_FORMAT_E As String = "#End If     '@#PACKAGEID#E"
Public Const TIL_FORMAT As String = "'@#PACKAGEID#  #DATE#   #NAME#" & vbTab & "#DESCRIPTION#"
Public Const REP_FORMAT_S As String = "#If 0 Then  '@#PACKAGEID#R"
Public Const REP_FORMAT_E As String = "#Else     '@#PACKAGEID#" & vbCrLf & vbCrLf & "#End If     '@#PACKAGEID#E"

Public Const PACKAGE_ID As String = "#PACKAGEID#"
Public Const DATETIME As String = "#DATE#"
Public Const AUTHOR As String = "#NAME#"
Public Const DESCRIPTION As String = "#DESCRIPTION#"
Public Const OTHER As String = "#OTHER#"
Public Const DESTDIR As String = "#DESTDIR#"
Public Const SORCDIR As String = "#SORCDIR#"

Public Const SECTION As String = "CIS_SET"
Public Const INI_FILE  As String = "CisSetting.ini"
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public gPackageId As String
Public gAuthor As String
Public gDateTime As String
Public gDescription As String
Public gOther As String
Public gDestDir As String
Public gSorcDir As String

Public gVbInstance  As VBIDE.VBE

Public gSearchedComponents As Collection
Public gSearchedTexts As Collection

Public cmdSearchText As Office.CommandBarButton
 
Public Function ADD_STRING_S() As String
    ADD_STRING_S = Replace(ADD_FORMAT_S, PACKAGE_ID, gPackageId)
End Function

Public Function ADD_STRING_E() As String
    ADD_STRING_E = Replace(ADD_FORMAT_E, PACKAGE_ID, gPackageId)
End Function

Public Function REP_STRING_S() As String
    REP_STRING_S = Replace(REP_FORMAT_S, PACKAGE_ID, gPackageId)
End Function
Public Function REP_STRING_E() As String
    REP_STRING_E = Replace(REP_FORMAT_E, PACKAGE_ID, gPackageId)
End Function

Public Function DEL_STRING_S() As String
    DEL_STRING_S = Replace(DEL_FORMAT_S, PACKAGE_ID, gPackageId)
End Function
Public Function DEL_STRING_E() As String
    DEL_STRING_E = Replace(DEL_FORMAT_E, PACKAGE_ID, gPackageId)
End Function
Public Function TIL_STRING() As String
    TIL_STRING = Replace(TIL_FORMAT, PACKAGE_ID, gPackageId)
    TIL_STRING = Replace(TIL_STRING, DATETIME, gDateTime)
    TIL_STRING = Replace(TIL_STRING, AUTHOR, Trim(gAuthor))
    TIL_STRING = Replace(TIL_STRING, DESCRIPTION, gDescription)
    
    If gOther <> Empty Then
        TIL_STRING = TIL_STRING & vbCrLf & gOther
    End If
End Function


Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
   As Long

   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
         0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
         0, 0, FLAGS)
      SetTopMostWindow = False
   End If
End Function

Public Function GetINI(sINIFile As String, sSection As String, sKey As String, sDefault As String) As String
    Dim sTemp As String * 255
    Dim nLength As Integer
    sTemp = Space$(256)
    nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sINIFile)
    nLength = InStr(1, sTemp, Chr(0), vbBinaryCompare)
    GetINI = Left$(sTemp, nLength - 1)
End Function

Public Sub WriteINI(sINIFile As String, sSection As String, sKey As String, sValue As String)
    Dim n As Integer
    Dim sTemp As String
    sTemp = sValue
    For n = 1 To Len(sValue)
        If Mid$(sValue, n, 1) = vbCr Or Mid$(sValue, n, 1) = vbLf Then Mid$(sValue, n) = " "
    Next n
    n = WritePrivateProfileString(sSection, sKey, sTemp, sINIFile)
End Sub

Public Sub WriteAll()
    Call WriteINI(INI_FILE, SECTION, PACKAGE_ID, gPackageId)
    Call WriteINI(INI_FILE, SECTION, DATETIME, gDateTime)
    Call WriteINI(INI_FILE, SECTION, AUTHOR, gAuthor)
    Call WriteINI(INI_FILE, SECTION, DESCRIPTION, gDescription)
    Call WriteINI(INI_FILE, SECTION, OTHER, gOther)
    Call WriteINI(INI_FILE, SECTION, DESTDIR, gDestDir)
    Call WriteINI(INI_FILE, SECTION, SORCDIR, gSorcDir)

End Sub

Public Sub ReadAll()
    gPackageId = GetINI(INI_FILE, SECTION, PACKAGE_ID, gPackageId)
    gDateTime = GetINI(INI_FILE, SECTION, DATETIME, gDateTime)
    gAuthor = GetINI(INI_FILE, SECTION, AUTHOR, gAuthor)
    gDescription = GetINI(INI_FILE, SECTION, DESCRIPTION, gDescription)
    gOther = GetINI(INI_FILE, SECTION, OTHER, gOther)
    gDestDir = GetINI(INI_FILE, SECTION, DESTDIR, "C:\")
    gSorcDir = GetINI(INI_FILE, SECTION, SORCDIR, "C:\")
End Sub

Public Sub DisplayHScroll(listbox As Object, Optional width As Integer = 1500)
    Dim Information#, Scrollbar#
   
   Information = SendMessage(listbox.hwnd, LB_SETHORIZONTALEXTENT, width, 0)
   Scrollbar = GetWindowLong(listbox.hwnd, GWL_STYLE)
   Scrollbar = Scrollbar Or WS_HSCROLL
   SetWindowLong listbox.hwnd, GWL_STYLE, Scrollbar
   SetWindowPos listbox.hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOSIZE Or SWP_FRAMECHANGED

End Sub

Public Function GetFolder(filename As String)
    Dim paths() As String, j As Integer
    paths = Split(filename, "\")
    For j = LBound(paths) To UBound(paths) - 1
        GetFolder = GetFolder & paths(j) & "\"
    Next
End Function

Public Function convertColumn2Pos(strText As String, column As Long) As Long
    Dim i As Integer
    Dim col As Integer
    col = 0
    For i = 1 To Len(strText)
        col = col + IIf(Asc(Mid(strText, i, 1)) < 0, 2, 1)
        If col >= column Then convertColumn2Pos = i: Exit For
    Next i
    If convertColumn2Pos = 0 Then convertColumn2Pos = i
End Function


