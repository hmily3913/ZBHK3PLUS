VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BillEvent_CJJGJS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is BillEvent Interface Class, made by K3BOSPLUGINSWIZAED"
 
'定义 BillEvent 接口. 必须具有的声明, 以此来获得事件
Private WithEvents m_BillInterface  As BillEvent
Attribute m_BillInterface.VB_VarHelpID = -1
 
Public Sub Show(ByVal oBillInterface As Object)
 
    'BillEvent 接口实现
    '注意: 此方法必须存在, 请勿修改
    Set m_BillInterface = oBillInterface
 
End Sub

Private Sub Class_Terminate()
 
    '释放接口对象
    '注意: 此方法必须存在, 请勿修改
    Set m_BillInterface = Nothing

End Sub

Private Sub m_BillInterface_Change(ByVal dct As KFO.IDictionary, ByVal dctFld As KFO.IDictionary, ByVal Col As Long, ByVal Row As Long, Cancel As Boolean)
 
    'TODO: 请在此处添加代码响应事件
    If Col <> -1 And Row <> -1 Then
    If UCase(dct.Value("FFieldName")) = UCase("FBase2") Or UCase(dct.Value("FFieldName")) = UCase("FDecimal") Then
    If m_BillInterface.GetFieldValue("FBase2", Row, Enu_ValueType_FFLD) <> "" And m_BillInterface.GetFieldValue("FDecimal", Row, Enu_ValueType_FFLD) <> "" Then
        On Error GoTo Catch
        Dim strSQL As String
        Dim rs As ADODB.Recordset
        Set oK3 = m_BillInterface.K3Lib
        Dim L, W, H, mj, price As Double
        L = m_BillInterface.GetFieldValue("FBaseProperty2", Row, Enu_ValueType_FFLD)
        W = m_BillInterface.GetFieldValue("FBaseProperty3", Row, Enu_ValueType_FFLD)
        H = m_BillInterface.GetFieldValue("FBaseProperty4", Row, Enu_ValueType_FFLD)
        strSQL = "select * from z_BCJSGSEntry where FID=" & m_BillInterface.GetFieldValue("FBase2", Row, Enu_ValueType_FFLD)
        Set rs = oK3.GetData(strSQL)
        Dim gs '公式
        gs = rs("FText")
        gs = Replace(Replace(Replace(gs, "长", L), "宽", W), "高", H)
        
        Dim objScript As Object
        Set objScript = CreateObject("MSScriptControl.ScriptControl")
        objScript.Language = "VBScript"
        mj = objScript.Eval(gs) / 1000 / 1000 '平方米
        price = mj * m_BillInterface.GetFieldValue("FDecimal", Row, Enu_ValueType_FFLD)
        
        Dim bln As Boolean
        bln = m_BillInterface.SetFieldValue("FDecimal1", price, Row)
        GoTo ExitSub
Catch:
        Set dct = Nothing
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbCritical, oK3.LoadKDString("金蝶提示")
        End If
ExitSub:
        Set oK3 = Nothing
        Set rs = Nothing
        Set dct = Nothing
    End If
    End If
    End If
End Sub


