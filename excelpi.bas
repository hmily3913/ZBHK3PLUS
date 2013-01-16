Attribute VB_Name = "excelpi"
'新建一个模块Module，复制如下代码到里面
Option Explicit
    
Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
          ByVal hWnd As Long, _
          ByVal wMsg As Long, _
          ByVal wParam As Long, _
          ByVal lParam As String) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" ( _
          ByVal pidl As Long, _
          ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" ( _
          lpBrowseInfo As BROWSEINFO) As Long
Type BROWSEINFO
          hOwner   As Long
          pidlRoot   As Long
          pszDisplayName   As String
          lpszTitle   As String
          ulFlags   As Long
          lpfnCallback   As Long
          lParam   As Long
          iImage   As Long
End Type
Dim xStartPath     As String
    
Function SelectDir(Optional StartPath As String, Optional Titel As String) As String
          Dim iBROWSEINFO     As BROWSEINFO
          With iBROWSEINFO
                  .lpszTitle = IIf(Len(Titel), Titel, "【请选择文件夹】")
                  .ulFlags = 7
                  If Len(StartPath) Then
                  xStartPath = StartPath & vbNullChar
                  .lpfnCallback = GetAddressOf(AddressOf CallBack)
                  End If
          End With
          Dim xPath     As String, NoErr       As Long:     xPath = Space$(512)
          NoErr = SHGetPathFromIDList(SHBrowseForFolder(iBROWSEINFO), xPath)
          SelectDir = IIf(NoErr, Left$(xPath, InStr(xPath, Chr(0)) - 1), "")
End Function
    
Function GetAddressOf(Address As Long) As Long
          GetAddressOf = Address
End Function
    
Function CallBack(ByVal hWnd As Long, _
                                      ByVal Msg As Long, _
                                      ByVal pidl As Long, _
                                      ByVal pData As Long) As Long
          Select Case Msg
                  Case 1
                          Call SendMessage(hWnd, 1126, 1, xStartPath)
                  Case 2
                          Dim sDir     As String * 64, tmp           As Long
                          tmp = SHGetPathFromIDList(pidl, sDir)
                          If tmp = 1 Then SendMessage hWnd, 1124, 0, sDir
          End Select
End Function
    

