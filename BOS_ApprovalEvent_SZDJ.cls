VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_ApprovalEvent_SZDJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is ApprovalEvents Interface Class, made by K3BOSPLUGINSWIZAED"
 

Private WithEvents m_ApprovalInterface  As MixedApprovalEvents
Attribute m_ApprovalInterface.VB_VarHelpID = -1
 
Public Sub Show(ByVal oApprovalInterface As Object)
 
    
    'ע��: �˷����������, �����޸�
    Set m_ApprovalInterface = oApprovalInterface
 
End Sub

Private Sub Class_Terminate()
 
    '�ͷŽӿڶ���
    'ע��: �˷����������, �����޸�
    Set m_ApprovalInterface = Nothing

End Sub

Private Sub m_ApprovalInterface_CheckDone(ByVal lCheckNodeIndex As Long, ByVal bSucceed As Variant)
 
    'TODO: ���ڴ˴����Ӵ�����Ӧ�¼�
    On Error GoTo Catch
    Dim oK3  As K3ClassEvents.K3Lib
    Dim oUser As K3ClassEvents.User
    Dim oUserid As Long
    Dim strSQL As String
    Dim vctRows As KFO.Vector
    Dim dctRow As KFO.Dictionary
    Dim rs As New ADODB.Recordset
    
    If m_ApprovalInterface.BillEvent Is Nothing Then '��ʱ��������ִ��
        Set m_CZEvent = m_ApprovalInterface.ListEvent
        Set oK3 = m_CZEvent.K3Lib
        Set oUser = oK3.User
        oUserid = oUser.UserID
        Set vctRows = m_CZEvent.GetSelectedBillInfo
        If Not vctRows Is Nothing And bSucceed = True Then
            'ѭ����ʱ��
            For i = 1 To vctRows.Size
                Set dctRow = vctRows(i)
                strSQL = "select FTagIndex from icClassMCStatus200000037 WHERE FID =(SELECT MAX(FID) FROM icClassMCStatus200000037 WHERE FBillID=" & dctRow("FID") & " AND FTagIndex NOT IN(-1))"
                Set rs = oK3.GetData(strSQL)
                If Not rs.EOF Then
                    If rs("FTagIndex") = 2000 And lCheckNodeIndex = 2000 Then
                        strSQL = "select * from z_SZDJ where FID='" & dctRow("FID") & "'  "
                        Set rs = oK3.GetData(strSQL)
                        While Not rs.EOF
                            If rs("FComboBox") = "ũҩ������׼֤" Then
                                oK3.UpdateData ("Update t_ICItem set F_117='" & rs("FText5") & "',F_118='" & rs("FDate1") & "'  where FItemid=" & rs("FBase"))
                            ElseIf rs("FComboBox") = "ũҩ��׼" Then
                                oK3.UpdateData ("Update t_ICItem set F_119='" & rs("FText5") & "',F_120='" & rs("FDate1") & "'  where FItemid=" & rs("FBase"))
                            ElseIf rs("FComboBox") = "ũҩ�Ǽ�֤" Then
                                oK3.UpdateData ("Update t_ICItem set F_121='" & rs("FText5") & "',F_122='" & rs("FDate1") & "'  where FItemid=" & rs("FBase"))
                            End If
                            rs.MoveNext
                        Wend
                    End If
                End If
            Next
'            m_CZEvent.RefreshList True, True
            Set rs = Nothing
            Set dctRow = Nothing
            Set vctRows = Nothing
        End If
    Else '���ݽ�����ִ��
        Set m_CZEvent = m_ApprovalInterface.BillEvent
        Set oK3 = m_CZEvent.K3Lib
        Set oUser = oK3.User
        oUserid = oUser.UserID
        
        If m_CZEvent.Data("page1")(m_CZEvent.TableInfo("map")("FID"))("FFLD") <> "" Then
            If lCheckNodeIndex = 2000 And bSucceed = True Then
                strSQL = "select * from z_SZDJ where FID='" & m_CZEvent.Data("page1")(m_CZEvent.TableInfo("map")("FID"))("FFLD") & "'  "
                Set rs = oK3.GetData(strSQL)
                While Not rs.EOF
                    If rs("FComboBox") = "ũҩ������׼֤��" Then
                        oK3.UpdateData ("Update t_ICItem set F_117='" & rs("FText5") & "',F_118='" & rs("FDate1") & "'  where FItemid=" & rs("FBase"))
                    ElseIf rs("FComboBox") = "ũҩ��׼" Then
                        oK3.UpdateData ("Update t_ICItem set F_119='" & rs("FText5") & "',F_120='" & rs("FDate1") & "'  where FItemid=" & rs("FBase"))
                    ElseIf rs("FComboBox") = "ũҩ�Ǽ�֤" Then
                        oK3.UpdateData ("Update t_ICItem set F_121='" & rs("FText5") & "',F_122='" & rs("FDate1") & "'  where FItemid=" & rs("FBase"))
                    End If
                    rs.MoveNext
                Wend
                rs.Close
                Set rs = Nothing
            End If
        End If
    End If
    GoTo ExitSub
Catch:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, oK3.LoadKDString("�����ʾ")
    End If
ExitSub:
    Set oUser = Nothing
    Set oK3 = Nothing
End Sub

