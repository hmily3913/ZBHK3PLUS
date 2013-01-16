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
 
'���� BillEvent �ӿ�. ������е�����, �Դ�������¼�
Private WithEvents m_BillInterface  As BillEvent
Attribute m_BillInterface.VB_VarHelpID = -1
 
Public Sub Show(ByVal oBillInterface As Object)
 
    'BillEvent �ӿ�ʵ��
    'ע��: �˷����������, �����޸�
    Set m_BillInterface = oBillInterface
 
End Sub

Private Sub Class_Terminate()
 
    '�ͷŽӿڶ���
    'ע��: �˷����������, �����޸�
    Set m_BillInterface = Nothing

End Sub

Private Sub m_BillInterface_Change(ByVal dct As KFO.IDictionary, ByVal dctFld As KFO.IDictionary, ByVal Col As Long, ByVal Row As Long, Cancel As Boolean)
 
    'TODO: ���ڴ˴����Ӵ�����Ӧ�¼�
    If UCase(dct.Value("FFieldName")) = UCase("FBillType") Then
'        Dim Accountfld As BOSField
'        Set Accountfld = m_BillInterface.BillEntrys(2).BOSFields(37)
'        Accountfld.Row = 1
'        Accountfld.Value = 1388
'        m_BillInterface.Data("page3")(1)(m_BillInterface.TableInfo("page3")(37))("FFLD") = 1388

        If dctFld("FDSP") = "����" Then
        m_BillInterface.SetFieldValue "FExplanation", "����", 1
        m_BillInterface.SetFieldValue "FEntryAccountID", "2181.14", 1
        ElseIf dctFld("FDSP") = "���˽�" Then
        m_BillInterface.SetFieldValue "FExplanation", "���˽�", 1
        m_BillInterface.SetFieldValue "FEntryAccountID", "2181.06", 1
        ElseIf dctFld("FDSP") = "�˷�" Then
        m_BillInterface.SetFieldValue "FExplanation", "�˷�", 1
        m_BillInterface.SetFieldValue "FEntryAccountID", "2181.10", 1
        ElseIf dctFld("FDSP") = "Ʒ�ƴ�����" Then
        m_BillInterface.SetFieldValue "FExplanation", "Ʒ�ƴ�����", 1
        m_BillInterface.SetFieldValue "FEntryAccountID", "2181.08", 1
        ElseIf dctFld("FDSP") = "������֤��" Then
        m_BillInterface.SetFieldValue "FExplanation", "������֤��", 1
        m_BillInterface.SetFieldValue "FEntryAccountID", "2181.13", 1
        End If
    End If
End Sub
