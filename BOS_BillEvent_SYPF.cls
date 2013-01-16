VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BillEvent_SYPF"
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
    On Error GoTo Catch
    
    Dim strSQL As String
    Dim oK3  As K3ClassEvents.K3Lib
    Dim rs As ADODB.Recordset
    Set oK3 = m_BillInterface.K3Lib
    Dim bln As Boolean
    If UCase(dct.Value("FFieldName")) = UCase("FBase1") Then
        
        strSQL = " select a.FName,a.FUnitID,a.FSecUnitID,b.FNumber as Num1,c.FNumber as Num2 from t_icitem a left join t_measureUnit b on a.FUnitID=b.Fitemid left join t_measureUnit c on a.FSecUnitID=c.FItemID where a.fitemid= " & m_BillInterface.GetFieldValue("FBase1", Row, Enu_ValueType_FFLD)
        Set rs = oK3.GetData(strSQL)
        m_BillInterface.SetFieldValue "FText1", rs("FName"), Row
        m_BillInterface.SetFieldValue "FBase5", rs("Num1"), Row
        If rs("FSecUnitID") = 0 Then
        m_BillInterface.SetFieldValue "FComboBox1", rs("FUnitID"), Row
        Else
        m_BillInterface.SetFieldValue "FComboBox1", rs("FSecUnitID"), Row
        End If
    End If
    
    GoTo ExitSub
Catch:
    Set dct = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, oK3.LoadKDString("金蝶提示")
    End If
ExitSub:
    Set oK3 = Nothing
    Set rs = Nothing
End Sub

Private Sub m_BillInterface_CommandClick(ByVal CommandName As String, bCancel As Boolean)
 
    'TODO: 请在此处添加代码响应事件
    '响应计算按钮
    If CommandName = "FButton" Then
    Dim lSum, tianSum As Double
    lSum = 0
    For i = 1 To m_BillInterface.Data("Page2").Size - 1
        lSum = lSum + m_BillInterface.GetFieldValue("FDecimal5", i, Enu_ValueType_FFLD)
    Next
    tianSum = m_BillInterface.GetFieldValue("FDecimal1", , Enu_ValueType_FFLD) / 1000 - lSum
    m_BillInterface.SetFieldValue "FDecimal5", tianSum, m_BillInterface.Data("Page2").Size
    Dim blnResult As Boolean
    blnResult = m_BillInterface.BatchToDoActions("FDecimal5", m_BillInterface.Data("Page2").Size, m_BillInterface.Data("Page2").Size)

    End If
End Sub
