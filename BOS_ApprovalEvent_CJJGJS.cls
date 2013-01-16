VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_ApprovalEvent_CJJGJS"
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
            strSQL = "select FTagIndex from icClassMCStatus200000035 WHERE FID =(SELECT MAX(FID) FROM icClassMCStatus200000035 WHERE FBillID=" & dctRow("FID") & " AND FTagIndex NOT IN(-1))"
            Set rs = oK3.GetData(strSQL)
            If Not rs.EOF Then
                If rs("FTagIndex") = 2000 And lCheckNodeIndex = 2000 Then
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
            If lCheckNodeIndex = 2000 And bSucceed = True Then
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
    
    Set cn1 = New ADODB.Connection
    cn1.CommandTimeout = 0
    cn1.Open "Driver=Sql server;server=10.0.0.121;pwd=mflogin;Uid=sa;DataBase=AIS20091116143745"
    
    strSQL = "select B.*,A.FComboBox from z_CJJGJSEntry b,z_CJJGJS a where a.FID=B.FID AND a.FID='" & FID & "'  "
    Set rs = oK3.GetData(strSQL)
    While Not rs.EOF
    '开始循环表体
        '如果是采购报价
        If rs("FCOMBOBOX") = "01" Then
        strSql2 = "SELECT 1 FROM t_SupplyEntry WHERE FItemID=" & rs("FBase") & " AND FSupID=" & rs("FBase1") & " AND FUnitID=" & rs("FBase3") & " AND FStartQty=0 AND FEndQty=0 AND FPType=1 AND FCyID=1 AND FQuoteTime=getdate() AND FDisableDate='2100-01-01' "
        Set rs2 = oK3.GetData(strSql2)
        If Not rs2.EOF Then
            MsgBox "分录为" & rs("Fentryid") & "的采购价格单据已存在，添加失败！", vbCritical, oK3.LoadKDString("金蝶提示")
            GoTo ExitSub
        End If
        '获取内码
        Dim newFID
        Dim adoComm, prm
        Set adoComm = CreateObject("ADODB.Command")
        With adoComm
        .ActiveConnection = cn1
        .CommandType = 4
        .CommandText = "GetICMaxNum"
        Set prm = .CreateParameter("TableName", 200, 1, 50, "t_Supply")
        .Parameters.Append prm
        Set prm = .CreateParameter("FInterID", 3, 2, , 1)
        .Parameters.Append prm
        .Execute
        newFID = .Parameters.Item("FInterID").Value
        End With
        Set adoComm = Nothing
        
        '取当前单据编号序列号
        Dim FBillNo, fprojectval
        strSql2 = "select FBrNo FROM t_Supply WHERE FItemID=" & rs("FBase") & " AND FSupID=" & rs("FBase1") & " AND FPType=1  "
        Set rs2 = oK3.GetData(strSql2)
        If rs2.EOF Then
            oK3.GetData ("INSERT INTO t_Supply (FBrNo,FItemID,FSupID,FCurrencyID,FPOHighPrice,FPType) VALUES ('0'," & rs("FBase") & "," & rs("FBase1") & ",1,0,1)")
            strSql2 = "INSERT INTO t_SupplyEntry (FBrNo,FUsed,FEntryID,FSupID,FItemID,FUnitID,FStartQty,FEndQty,FPType,FPrice,FCyID,FDisCount,FLeadTime,FQuoteTime,FDisableDate,FRemark, FLastModifiedBy, FLastModifiedDate) "
            strSql2 = strSql2 + " VALUES ('0',0," & newFID & "," & rs("FBase1") & "," & rs("FBase") & "," & rs("FBase3") & ",0,0,1," & rs("FDecimal1") & ",1,0.000000,0,getdate(),'2100-01-01',''," & oUserid & ",getdate())"
            oK3.GetData (strSql2)
        Else
            strSql2 = "INSERT INTO t_SupplyEntry (FBrNo,FUsed,FEntryID,FSupID,FItemID,FUnitID,FStartQty,FEndQty,FPType,FPrice,FCyID,FDisCount,FLeadTime,FQuoteTime,FDisableDate,FRemark, FLastModifiedBy, FLastModifiedDate) "
            strSql2 = strSql2 + " VALUES ('0',0," & newFID & "," & rs("FBase1") & "," & rs("FBase") & "," & rs("FBase3") & ",0,0,1," & rs("FDecimal1") & ",1,0.000000,0,getdate(),'2100-01-01',''," & oUserid & ",getdate())"
            oK3.GetData (strSql2)
        End If
        ElseIf rs("FCOMBOBOX") = "02" Then
            strSql2 = "select 1 from t_ZJHDJDR a,t_ZJHDJDREntry b where a.fid=b.fid and a.fuser=0 and a.FClassTypeid='" & rs("FClassID_Src") & "' and b.FID=" & rs("FID_Src") & " and b.FEntryID=" & rs("FEntryID_Src") & " and b.fbase=" & rs("FBase")
            Set rs2 = oK3.GetData(strSql2)
            If Not rs2.EOF Then
                strSql2 = "update t_ZJHDJDREntry set FDecimal=" & rs("FDecimal1") & " where FID=" & rs("FID_Src") & " and FEntryID=" & rs("FEntryID_Src")
                oK3.GetData (strSql2)
            End If
        End If
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

