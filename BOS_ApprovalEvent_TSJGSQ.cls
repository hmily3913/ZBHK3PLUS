VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_ApprovalEvent_TSJGSQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is ApprovalEvents Interface Class, made by K3BOSPLUGINSWIZAED"
 

Private WithEvents m_ApprovalInterface  As MixedApprovalEvents
Attribute m_ApprovalInterface.VB_VarHelpID = -1
Dim m_CZEvent As Object
 
Public Sub Show(ByVal oApprovalInterface As Object)
 
    
    '注意: 此方法必须存在, 请勿修改
    Set m_ApprovalInterface = oApprovalInterface
 
End Sub

Private Sub Class_Terminate()
 
    '释放接口对象
    '注意: 此方法必须存在, 请勿修改
    Set m_ApprovalInterface = Nothing

End Sub

Private Sub m_ApprovalInterface_CheckDone(ByVal lCheckNodeIndex As Long, ByVal bSucceed As Variant)
 
    'TODO: 请在此处添加代码响应事件
    On Error GoTo Catch
    Dim oK3  As K3ClassEvents.K3Lib
    Dim oUser As K3ClassEvents.User
    Dim oUserid As Long
    Dim strSQL As String
    Dim vctRows As KFO.Vector
    Dim dctRow As KFO.Dictionary
    Dim rs As New ADODB.Recordset
    
    If m_ApprovalInterface.BillEvent Is Nothing Then '序时簿界面中执行
        Set m_CZEvent = m_ApprovalInterface.ListEvent
        Set oK3 = m_CZEvent.K3Lib
        Set oUser = oK3.User
        oUserid = oUser.UserID
        Set vctRows = m_CZEvent.GetSelectedBillInfo
        If Not vctRows Is Nothing And bSucceed = True Then
            '循环序时簿
            For i = 1 To vctRows.Size
                Set dctRow = vctRows(i)
                strSQL = "select FTagIndex from icClassMCStatus200000023 WHERE FID =(SELECT MAX(FID) FROM icClassMCStatus200000023 WHERE FBillID=" & dctRow("FID") & " AND FTagIndex NOT IN(-1))"
                Set rs = oK3.GetData(strSQL)
                If Not rs.EOF Then
                    If rs("FTagIndex") = 2002 And lCheckNodeIndex = 2002 Then
                        strSQL = "update  t_BOS200000023  Set FUser1=" & oUserid & ",FTime2=getdate() where FID='" & dctRow("FID") & "'  "
                        oK3.GetData (strSQL)
                    ElseIf rs("FTagIndex") = 2001 And lCheckNodeIndex = 2001 Then
                        strSQL = "update  t_BOS200000023  Set FUser=" & oUserid & ",FTime1=getdate() where FID='" & dctRow("FID") & "'  "
                        oK3.GetData (strSQL)
                    End If
                End If
            Next
'            m_CZEvent.RefreshList True, True
            Set rs = Nothing
            Set dctRow = Nothing
            Set vctRows = Nothing
        End If
    Else '单据界面中执行
        Set m_CZEvent = m_ApprovalInterface.BillEvent
        Set oK3 = m_CZEvent.K3Lib
        Set oUser = oK3.User
        oUserid = oUser.UserID
        
        If m_CZEvent.Data("page1")(m_CZEvent.TableInfo("map")("FID"))("FFLD") <> "" Then
            If lCheckNodeIndex = 2001 And bSucceed = True Then
                strSQL = "  update  t_BOS200000023  Set FUser=" & oUserid & ",FTime1=getdate() where FID='" & m_CZEvent.Data("page1")(m_CZEvent.TableInfo("map")("FID"))("FFLD") & "'  "
                Set rs = oK3.GetData(strSQL)
                m_CZEvent.SetFieldValue "FUser", oUserid
                m_CZEvent.SetFieldValue "FTime1", Now()
            ElseIf lCheckNodeIndex = 2002 And bSucceed = True Then
                strSQL = "  update  t_BOS200000023  Set FUser1=" & oUserid & ",FTime2=getdate() where FID='" & m_CZEvent.Data("page1")(m_CZEvent.TableInfo("map")("FID"))("FFLD") & "'  "
                Set rs = oK3.GetData(strSQL)
                m_CZEvent.SetFieldValue "FUser1", oUserid
                m_CZEvent.SetFieldValue "FTime2", Now()
            End If
        End If
    End If
    GoTo ExitSub
Catch:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, oK3.LoadKDString("金蝶提示")
    End If
ExitSub:
    Set oUser = Nothing
    Set oK3 = Nothing
End Sub

Private Sub m_ApprovalInterface_UnCheckDone(ByVal lCheckNodeIndex As Long, ByVal bSucceed As Variant)
 
    'TODO: 请在此处添加代码响应事件
    On Error GoTo Catch
    Dim oK3  As K3ClassEvents.K3Lib
    Dim oUser As K3ClassEvents.User
    Dim oUserid As Long
    Dim strSQL As String
    Dim vctRows As KFO.Vector
    Dim dctRow As KFO.Dictionary
    Dim rs As New ADODB.Recordset
    
    If m_ApprovalInterface.BillEvent Is Nothing Then '序时簿界面中执行
        Set m_CZEvent = m_ApprovalInterface.ListEvent
        Set oK3 = m_CZEvent.K3Lib
        Set oUser = oK3.User
        oUserid = oUser.UserID
        Set vctRows = m_CZEvent.GetSelectedBillInfo
        If Not vctRows Is Nothing And bSucceed = True Then
            For i = 1 To vctRows.Size
                Set dctRow = vctRows(i)
                strSQL = "select FNextNodeTagIndex from icClassMCStatus200000023 WHERE FID =(SELECT MAX(FID) FROM icClassMCStatus200000023 WHERE FBillID=" & dctRow("FID") & " AND FTagIndex NOT IN(-1))"
                Set rs = oK3.GetData(strSQL)
                If Not rs.EOF Then
                    If rs("FNextNodeTagIndex") = 2002 Then
                        strSQL = "  update  t_BOS200000023  Set FUser1=0,FTime2=null where FID='" & dctRow("FID") & "'  "
                        oK3.GetData (strSQL)
                    ElseIf rs("FNextNodeTagIndex") = 2001 Then
                        strSQL = "  update  t_BOS200000023  Set FUser=0,FTime1=null where FID='" & dctRow("FID") & "'  "
                        oK3.GetData (strSQL)
                    End If
                End If
            Next
            m_CZEvent.RefreshList True, True
            Set rs = Nothing
            Set dctRow = Nothing
            Set vctRows = Nothing
        End If
    Else '单据界面中执行
        Set m_CZEvent = m_ApprovalInterface.BillEvent
        Set oK3 = m_CZEvent.K3Lib
        Set oUser = oK3.User
        oUserid = oUser.UserID
        
        If m_CZEvent.Data("page1")(m_CZEvent.TableInfo("map")("FID"))("FFLD") <> "" And bSucceed = True Then
            strSQL = "select FNextNodeTagIndex from icClassMCStatus200000023 WHERE FID =(SELECT MAX(FID) FROM icClassMCStatus200000023 WHERE FBillID=" & m_CZEvent.Data("page1")(m_CZEvent.TableInfo("map")("FID"))("FFLD") & " AND FTagIndex NOT IN(-1))"
            Set rs = oK3.GetData(strSQL)
            If Not rs.EOF Then
                If rs("FNextNodeTagIndex") = 2002 Then
                    strSQL = "  update  t_BOS200000023  Set FUser1=null,FTime2=null where FID='" & dctRow("FID") & "'  "
                    oK3.GetData (strSQL)
                    m_CZEvent.SetFieldValue "FUser1", 0
                    m_CZEvent.SetFieldValue "FTime2", Null
                ElseIf rs("FNextNodeTagIndex") = 2001 Then
                    strSQL = "  update  t_BOS200000023  Set FUser=null,FTime1=null where FID='" & dctRow("FID") & "'  "
                    oK3.GetData (strSQL)
                    m_CZEvent.SetFieldValue "FUser", 0
                    m_CZEvent.SetFieldValue "FTime1", Null
                End If
            End If
        End If
    End If
    GoTo ExitSub
Catch:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, oK3.LoadKDString("金蝶提示")
    End If
ExitSub:
    Set oUser = Nothing
    Set oK3 = Nothing
End Sub

