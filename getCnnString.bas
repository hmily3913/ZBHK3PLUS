Attribute VB_Name = "getCnnString"
Private Declare Function GetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessID" () As Long

'��ȡ���Ӵ�
Public Function GetPropsString() As String
    Dim lProc As Long
    Dim spmMgr As Object
    
    lProc = GetCurrentProcessId
    Set spmMgr = CreateObject("PropsMgr.ShareProps")
    GetPropsString = spmMgr.getproperty(lProc, "PropsString")
    
End Function
