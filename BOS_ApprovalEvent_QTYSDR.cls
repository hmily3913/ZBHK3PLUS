VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_ApprovalEvent_QTYSDR"
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
            strSQL = "select FTagIndex from icClassMCStatus200000031 WHERE FID =(SELECT MAX(FID) FROM icClassMCStatus200000031 WHERE FBillID=" & dctRow("FID") & " AND FTagIndex NOT IN(-1))"
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
    
    strSQL = "select * from z_QTYSDEntry where FID='" & FID & "'  "
    Set rs = oK3.GetData(strSQL)
    While Not rs.EOF
    '开始循环表体
        '获取内码
        Dim newFID
        Dim adoComm, prm
        Set adoComm = CreateObject("ADODB.Command")
        With adoComm
        .ActiveConnection = cn1
        .CommandType = 4
        .CommandText = "GetICMaxNum"
        Set prm = .CreateParameter("TableName", 200, 1, 50, "t_RP_ARPBill")
        .Parameters.Append prm
        Set prm = .CreateParameter("FInterID", 3, 2, , 1)
        .Parameters.Append prm
        .Execute
        newFID = .Parameters.Item("FInterID").Value
        End With
        Set adoComm = Nothing
        
        '取当前单据编号序列号
        Dim FBillNo, fprojectval
        strSql2 = "SELECT * FROM t_BillCodeRule WITH(READUNCOMMITTED) WHERE fbilltypeid='1000021' and fprojectid=3  "
        Set rs2 = oK3.GetData(strSql2)
        fprojectval = rs2("FProjectVal")
        FBillNo = "QTYS" & Right(rs2("FProjectVal") + 1000000, 6)
        strSql2 = "SELECT 1 FROM t_RP_ARPBill Where FNumber = '" & FBillNo & "' AND FBillID <> 0 AND FClassTypeID=1000021 "
        Set rs2 = oK3.GetData(strSql2)
        If Not rs2.EOF Then
            MsgBox "单据编号已存在，添加失败！", vbCritical, oK3.LoadKDString("金蝶提示")
            GoTo ExitSub
        End If
        '更新编号数据
        ReDim strSQLBatch(1)
        strSQLBatch(0) = "update t_billcoderule set fprojectval = fprojectval+1 where fbilltypeid='1000021' and fprojectid=3 "
        strSQLBatch(1) = "Update ICBillNo Set FCurNo = " & fprojectval & ",FDesc = 'QTYS+'+'" & Right(fprojectval + 1000001, 6) & "' where fbillid = 1000021"
        Call oK3.UpdateData(strSQLBatch)
        '插入对应表（其他应收主表，子表，收款计划，首付款明细）
        strSql2 = "select * from t_Organization where fitemid=" & rs("FBase")
        Set rs2 = oK3.GetData(strSql2)
        Dim department, employee
        department = rs2("FDepartment")
        employee = rs2("FEmployee")
        strSql2 = "INSERT INTO t_RP_ARPBill(FItemClassID,FBillType,FDate,FCustomer,FFincDate,FBillID,FExchangeRateType,FCurrencyID,FExchangeRate,FInterestRate,FAccountID,FExplanation,FDepartment,FEmployee,FContractNo,FNumber,FAmountFor,FAmount,FChecker,FPreparer,FTaskID,FResourceID,FOrderID,FRP,FRemainAmount,FRemainAmountFor,FClassTypeID,FRPDate,FAdjustExchangeRate,FPeriod,FYear,FContractID,FSource,FBudgetAmountFor,FAdjustAmount,FRPBank,FBankAcct,FBankAcctName,FSubSystemID,FObtainRateWay)"
        strSql2 = strSql2 + "Values(1,'995','" & rs("FDate") & "'," & rs("FBase") & ",'" & rs("FDate1") & "'," & newFID & ",1,1,1,0," & rs("FBase1") & ",'" & rs("FText") & "'," & department & "," & employee & ",'','" & FBillNo & "'," & rs("FDecimal") & "," & rs("FDecimal") & ",0," & oUserid & ",0,0,0,'1'," & rs("FDecimal") & "," & rs("FDecimal") & ",1000021,'" & rs("FDate") & "',1," & Month(rs("FDate1")) & "," & Year(rs("FDate1")) & ",0,0,0,0,'','','',0,0)"
        oK3.GetData (strSql2)
        strSql2 = "INSERT INTO t_RP_Plan_Ar(FEntryID,FBillID,FDate,FAmountFor,FRemainAmount,FRemainAmountFor,FOrgID,FAmount,FRP)"
        strSql2 = strSql2 + "Values(1," & newFID & ",'" & rs("FDate") & "'," & rs("FDecimal") & "," & rs("FDecimal") & "," & rs("FDecimal") & ",0," & rs("FDecimal") & ",1)"
        oK3.GetData (strSql2)
        strSql2 = "INSERT INTO t_rp_arpbillEntry(FEntryID,FLinkCheckAmount,FLinkCheckAmountFor,FBillID,FClassID_SRC,FBillNo_SRC,FContractNo,famountFor,FID_SRC,FEntryID_SRC,FAmount,FTaxRate,FTaxAmountFor,FTaxAmount,FAmountNoTaxFor,FAmountNoTax,FAmountFor_Commit,FRemainAmountFor,FRemainAmountFor_SRC,FAmount_Commit,FRemainAmount,FRemainAmount_SRC,FContractEntryID,FPayApplyAmountFor,FPayApplyAmount,FInvLinkCheckAmount,FInvLinkCheckAmountFor,FInvoiceAmount,FInvoiceAmountFor,FAPAcctID) "
        strSql2 = strSql2 + "Values(1,0,0," & newFID & ",0,'',''," & rs("FDecimal") & ",0,0," & rs("FDecimal") & ",0,0,0," & rs("FDecimal") & "," & rs("FDecimal") & ",0," & rs("FDecimal") & ",0,0," & rs("FDecimal") & ",0,0,0,0,0,0,0,0,0)"
        oK3.GetData (strSql2)
        strSql2 = "INSERT INTO t_RP_Contact (FYear,FPeriod,FRP,FType,FDate,FFincDate,FNumber,FCustomer,FDepartment,FEmployee,FCurrencyID,FExchangeRate,FAmount,FAmountFor,FRemainAmount,FRemainAmountFor,FRPBillID,FRPDate,FContractNo,FInterestRate,FBillType,FItemClassID,FExplanation,FPreparer)"
        strSql2 = strSql2 + "Values(" & Year(rs("FDate1")) & "," & Month(rs("FDate1")) & ",1,1,'" & rs("FDate") & "','" & rs("FDate1") & "','" & FBillNo & "'," & rs("FBase") & "," & department & "," & employee & ",1,1," & rs("FDecimal") & "," & rs("FDecimal") & "," & rs("FDecimal") & "," & rs("FDecimal") & "," & newFID & ",'" & rs("FDate") & "','',0,995,1,'" & rs("FText") & "'," & oUserid & ")"
        oK3.GetData (strSql2)
        strSql2 = "update t_organization  set  FLastTradeDate='" & rs("FDate") & "', FLastTradeAmount=" & rs("FDecimal") & " where FItemID=" & rs("FBase") & ""
        oK3.GetData (strSql2)
        strSql2 = "exec sp_executesql N'SELECT FID FROM t_RP_Contact WHERE FType=@P1 AND FNumber=@P2',N'@P1 smallint,@P2 varchar(255)',1,'" & FBillNo & "'"
        Set rs2 = oK3.GetData(strSql2)
        Dim SrcFID
        SrcFID = rs2("FID")
        strSql2 = "update t_rp_plan_ar set FOrgID=" & SrcFID & " where FBillID=" & newFID & " and FRP=1 and FIsInit=0"
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
