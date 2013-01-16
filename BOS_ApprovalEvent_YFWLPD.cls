VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_ApprovalEvent_YFWLPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is ApprovalEvents Interface Class, made by K3BOSPLUGINSWIZAED"
 

Private WithEvents m_ApprovalInterface  As MixedApprovalEvents
Attribute m_ApprovalInterface.VB_VarHelpID = -1
 
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
            If vctRows.Size <> 1 Then
                MsgBox "只能选择一张单据进行审核", vbCritical, oK3.LoadKDString("金蝶提示")
                GoTo ExitSub
            End If
            Set dctRow = vctRows(1)
            strSQL = "select FTagIndex from icClassMCStatus200000040 WHERE FID =(SELECT MAX(FID) FROM icClassMCStatus200000040 WHERE FBillID=" & dctRow("FID") & " AND FTagIndex NOT IN(-1))"
            Set rs = oK3.GetData(strSQL)
            If Not rs.EOF Then
                If rs("FTagIndex") = 2001 And lCheckNodeIndex = 2001 Then
                    POIn m_CZEvent, dctRow("FID")
                End If
            End If
            'm_CZEvent.RefreshList True, True
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
                POIn m_CZEvent, m_CZEvent.Data("page1")(m_CZEvent.TableInfo("map")("FID"))("FFLD")
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

Private Sub m_ApprovalInterface_UnCheckDone(ByVal lUnCheckNodeIndex As Long, ByVal bSucceed As Variant)
 
    'TODO: 请在此处添加代码响应事件
    

End Sub

Function POIn(ByVal evt As Variant, ByVal FID As Variant)

    On Error GoTo Catch
    Dim oK3  As K3ClassEvents.K3Lib
    Dim oUser As K3ClassEvents.User
    Dim oUserid As Long
    Dim strSQL As String
    Dim strSql2 As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Set oK3 = evt.K3Lib
    Set oUser = oK3.User
    oUserid = oUser.UserID
    
    Dim strSQLBatch() As String
    
    strSQL = "select B.* from z_YFWLPDEntry b,z_YFWLPD a where a.FID=B.FID AND a.FID='" & FID & "'  "
    Set rs = oK3.GetData(strSQL)
    While Not rs.EOF
    '开始循环表体
        '如果是采购报价
        strSql2 = "update z_KYWLEntry set FDecimal2=" & rs("FDecimal1") & ",FDecimal1=FDecimal*" & rs("FDecimal1") & " WHERE FID=" & rs("FBase")
        oK3.GetData (strSql2)
        rs.MoveNext
    Wend
    GoTo ExitSub
Catch:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, oK3.LoadKDString("金蝶提示")
    End If
ExitSub:
    Set oUser = Nothing
    Set oK3 = Nothing
End Function


