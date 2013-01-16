Attribute VB_Name = "selectFile"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Function SelectFiles()
    Dim CommonDialog1 As CommonDialog
    
        On Error GoTo ms:
        With CommonDialog1
            .Filter = "所有文件|*.*"
            .ShowOpen
        End With
        
        Dim strFileName  As String
        SelectFiles = CommonDialog1.FileName
'        If Len(strFileName) > 0 And Dir(strFileName) <> "" Then
'            ShellExecute Me.hWnd, "Open", strFileName, vbNullString, vbNullString, 1
'        End If
          
ms:
    MsgBox Err.Description
      
End Function
