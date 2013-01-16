VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BillEvent_SZDJ"
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
    If UCase(dct.Value("FFieldName")) = UCase("FText5") Then

        If Left(dctFld("FFLD"), 2) = "PD" Or Left(dctFld("FFLD"), 2) = "LS" Then
        m_BillInterface.SetFieldValue "FComboBox", "农药登记证"
        m_BillInterface.DoAction "FComboBox"
        ElseIf Left(dctFld("FFLD"), 2) = "HG" Or Left(dctFld("FFLD"), 2) = "GB" Then
        m_BillInterface.SetFieldValue "FComboBox", "农药标准"
        m_BillInterface.DoAction "FComboBox"
        m_BillInterface.SetFieldValue "FDate1", "2100-12-30"
        ElseIf Left(dctFld("FFLD"), 4) = "Q/MF" Then
        m_BillInterface.SetFieldValue "FComboBox", "农药标准"
        m_BillInterface.DoAction "FComboBox"
        ElseIf Left(dctFld("FFLD"), 2) = "XK" Or Left(dctFld("FFLD"), 3) = "HNP" Then
        m_BillInterface.SetFieldValue "FComboBox", "农药生产批准证"
        m_BillInterface.DoAction "FComboBox"
        End If
    ElseIf UCase(dct.Value("FFieldName")) = UCase("FBase") Then
        If Left(dctFld("FFND"), 5) = "E.S.1" Then
        m_BillInterface.SetFieldValue "FComboBox2", "除草剂"
        m_BillInterface.DoAction "FComboBox2"
        ElseIf Left(dctFld("FFND"), 5) = "E.S.2" Then
        m_BillInterface.SetFieldValue "FComboBox2", "杀虫剂"
        m_BillInterface.DoAction "FComboBox2"
        ElseIf Left(dctFld("FFND"), 5) = "E.S.3" Then
        m_BillInterface.SetFieldValue "FComboBox2", "杀菌剂"
        m_BillInterface.DoAction "FComboBox2"
        ElseIf Left(dctFld("FFND"), 5) = "E.S.4" Then
        m_BillInterface.SetFieldValue "FComboBox2", "植物生长调节剂"
        m_BillInterface.DoAction "FComboBox2"
        End If
    End If
End Sub

