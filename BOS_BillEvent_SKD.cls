VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BillEvent_SKD"
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
    If UCase(dct.Value("FFieldName")) = UCase("FBillType") Then
'        Dim Accountfld As BOSField
'        Set Accountfld = m_BillInterface.BillEntrys(2).BOSFields(37)
'        Accountfld.Row = 1
'        Accountfld.Value = 1388
'        m_BillInterface.Data("page3")(1)(m_BillInterface.TableInfo("page3")(37))("FFLD") = 1388

        If dctFld("FDSP") = "定金" Then
        m_BillInterface.SetFieldValue "FExplanation", "定金", 1
        m_BillInterface.SetFieldValue "FEntryAccountID", "2181.14", 1
        ElseIf dctFld("FDSP") = "加盟金" Then
        m_BillInterface.SetFieldValue "FExplanation", "加盟金", 1
        m_BillInterface.SetFieldValue "FEntryAccountID", "2181.06", 1
        ElseIf dctFld("FDSP") = "运费" Then
        m_BillInterface.SetFieldValue "FExplanation", "运费", 1
        m_BillInterface.SetFieldValue "FEntryAccountID", "2181.10", 1
        ElseIf dctFld("FDSP") = "品牌代理费" Then
        m_BillInterface.SetFieldValue "FExplanation", "品牌代理费", 1
        m_BillInterface.SetFieldValue "FEntryAccountID", "2181.08", 1
        ElseIf dctFld("FDSP") = "储备保证金" Then
        m_BillInterface.SetFieldValue "FExplanation", "储备保证金", 1
        m_BillInterface.SetFieldValue "FEntryAccountID", "2181.13", 1
        End If
    End If
End Sub

