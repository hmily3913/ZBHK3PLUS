VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BillEvent_ZDCPJS"
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

Private Sub m_BillInterface_CommandClick(ByVal CommandName As String, bCancel As Boolean)
 
    'TODO: 请在此处添加代码响应事件
    '响应计算按钮
    If CommandName = "FButton" Then
    On Error GoTo Catch
    
    Dim strSQL As String
    Dim strSQLBatch() As String
'    Dim i As Integer
    Dim oK3  As K3ClassEvents.K3Lib
    Dim rs As ADODB.Recordset
    Set oK3 = m_BillInterface.K3Lib
    
    If m_BillInterface.GetFieldValue("FComboBox", , Enu_ValueType_FFLD) = "" Or m_BillInterface.GetFieldValue("FBase", , Enu_ValueType_FFLD) = "" Then
        MsgBox "年份、客户不能为空，读取失败！", vbCritical, oK3.LoadKDString("金蝶提示")
        Exit Sub
    End If
    If m_BillInterface.GetFieldValue("FUser", , Enu_ValueType_FFLD) <> "" Then
        MsgBox "审核状态不允许进行此操作！", vbCritical, oK3.LoadKDString("金蝶提示")
        Exit Sub
    End If
    strSQL = "select 1 from z_ZDCPJL where '" & m_BillInterface.GetFieldValue("FComboBox", , Enu_ValueType_FFLD) & "'=FComboBox and FBase=" & m_BillInterface.GetFieldValue("FBase", , Enu_ValueType_FFLD) & " and FID<>" & m_BillInterface.CurBillID
    Set rs = oK3.GetData(strSQL)
    If Not rs.EOF Then
        MsgBox "该客户在该年内已经做过结算单，不允许重复生成！", vbCritical, oK3.LoadKDString("金蝶提示")
        Exit Sub
    End If
    '删除现在分录
    Dim bln As Boolean
    bln = m_BillInterface.DeleteEntryData(2)
    Dim ythis, ylast As Long
    ythis = m_BillInterface.GetFieldValue("FComboBox", , Enu_ValueType_FFLD)
    ylast = ythis - 1
      strSQL = "select t3.fitemid,t3.fnumber,t3.fname,t3.fmodel,t4.fname,allsl1,allsl2,  "
      strSQL = strSQL & " case when FBase2=14563 then FPrice_Base*t5.FCoefficient  "
      strSQL = strSQL & " when allsl2<>0 then a2/allsl2*t5.FCoefficient  "
      strSQL = strSQL & " when FBase3=14563 and allsl1>0 and a1>0 then a1/allsl1*t5.FCoefficient  "
      strSQL = strSQL & " when allsl1>0 then FDecimal*t5.FCoefficient  "
      strSQL = strSQL & " else 0 end as auxprice,  "
      strSQL = strSQL & " allsl1/t5.FCoefficient as auxsl1,allsl2/t5.FCoefficient as auxsl2,  "
      strSQL = strSQL & " case when FBase2=14563 then FPrice_Base "
      strSQL = strSQL & " when allsl2<>0 then a2/allsl2 "
      strSQL = strSQL & " when FBase3=14563 and allsl1>0 and a1>0 then a1/allsl1 "
      strSQL = strSQL & " when allsl1>0 then FDecimal "
      strSQL = strSQL & " else 0 end as price,  "
      strSQL = strSQL & " case when FBase2=14563 then FPrice_Base*allsl1  "
      strSQL = strSQL & " when allsl2<>0 and allsl1>0 then a2/allsl2*allsl1  "
      strSQL = strSQL & " when FBase3=14563 and allsl1>0 then a1  "
      strSQL = strSQL & " when allsl1>0 then FDecimal*allsl1  "
      strSQL = strSQL & " else 0 end as sale1,  "
      strSQL = strSQL & " case when FBase2=14563 then FPrice_Base*allsl2  "
      strSQL = strSQL & " when allsl2<>0 then a2  "
      strSQL = strSQL & " else 0 end as sale2 "
      strSQL = strSQL & " from (select t6.FBASE,v1.FCustID,t6.FPrice_Base,t6.FBase2,t6.FBase3,t6.FDecimal,  "
      strSQL = strSQL & " sum(case when year(v1.fdate)='" & ylast & "' then v2.FQty else 0 end) as allsl1,  "
      strSQL = strSQL & " sum(case when year(v1.fdate)='" & ythis & "' then v2.FQty else 0 end) as allsl2,  "
      strSQL = strSQL & " sum(case when v2.FQty>0 and year(v1.fdate)='" & ylast & "' then v2.FQty else 0 end) as zhengsl1,  "
      strSQL = strSQL & " sum(case when v2.FQty>0 and year(v1.fdate)='" & ythis & "' then v2.FQty else 0 end) as zhengsl2,  "
      strSQL = strSQL & " sum(case when year(v1.fdate)='" & ylast & "' then ISNULL(v2.FAmountincludetax,0)+ISNULL(v2.FENTRYSELFI0463,0)+isnull(t7.FEntrySelfB0163,0) else 0 end) as a1,  "
      strSQL = strSQL & " sum(case when year(v1.fdate)='" & ylast & "' then isnull(t7.FEntrySelfB0163,0) else 0 end) as yf" & ylast & ",  "
      strSQL = strSQL & " sum(case when year(v1.fdate)='" & ythis & "' then ISNULL(v2.FAmountincludetax,0)+ISNULL(v2.FENTRYSELFI0463,0)+isnull(t7.FEntrySelfB0163,0) else 0 end) as a2  "
      strSQL = strSQL & " from ICSale v1 inner Join ICSaleEntry v2 On v1.FInterID=v2.FInterID and v1.FCustID=" & m_BillInterface.GetFieldValue("FBase", , Enu_ValueType_FFLD)
      strSQL = strSQL & " inner join (select t4.FBASE,t4.FPrice_Base,T4.FBase2,t4.FDecimal,t4.FBase3,t6.fitemid  "
      strSQL = strSQL & " from z_ZDCP t3 inner join z_ZDCPENTRY t4 on t3.fid=t4.fid and t3.FComboBox=" & ythis & "  "
      strSQL = strSQL & " inner join t_icitem t5 on t4.FBASE=t5.fitemid  "
      strSQL = strSQL & " inner join t_icitem t6 on left(t6.fnumber,13)=left(t5.fnumber,13) and len(t6.fnumber)=14) t6  "
      strSQL = strSQL & " on v2.fitemid=t6.fitemid  "
      strSQL = strSQL & " left join ICStockBillEntry t7 on v2.FSourceInterID=t7.FInterID and v2.FSourceEntryID=t7.FEntryID  "
      strSQL = strSQL & " where (v1.fdate>='" & ythis & "-01-01' and v1.fdate<='" & ythis & "-09-30'   "
      strSQL = strSQL & " or v1.fdate>='" & ylast & "-01-01' and v1.fdate<='" & ylast & "-09-30' )  "
      strSQL = strSQL & " and v1.FTRANTYPE=80 AND v1.FCheckerID<>0 AND v1.FCANCELLATION=0  "
      strSQL = strSQL & " group by t6.FBASE,v1.FCustID,t6.FPrice_Base,t6.FBase2,t6.FDecimal,t6.FBase3 ) t1,t_icitem t3,t_measureUnit t4,t_measureUnit t5 ,t_Organization t2 "
      strSQL = strSQL & " where t1.FBASE=t3.fitemid and t3.funitid=t4.fitemid and t3.FUnitGroupid=t5.FUnitGroupid and t5.FItemid=t3.FSaleUnitID and t1.fcustid=t2.fitemid "
      strSQL = strSQL & " order by t2.fnumber,t3.fnumber "
     Set rs = oK3.GetData(strSQL)
    
    Dim i As Long
    i = 1
    Dim allsum As Double
    allsum = 0
    Set dct = m_BillInterface.TableInfo(m_BillInterface.TableInfo("Map")("FBase1"))
    While (Not rs.EOF)
        m_BillInterface.InsertNewRowAndFill 2, i, "FBase1", rs("FNUMBER"), "FDecimal", rs("allsl1"), "FDecimal1", rs("allsl2"), "FDecimal2", rs("price"), "FDecimal3", rs("sale1"), "FDecimal4", rs("sale2"), "FDecimal5", rs("sale2") - rs("sale1"), "FDecimal7", rs("auxsl1"), "FDecimal8", rs("auxsl2"), "FDecimal9", rs("auxprice"), "FText2", "是"
'        m_BillInterface.LockCell dct, i, True
        allsum = allsum + rs("sale2") - rs("sale1")
        i = i + 1
        rs.MoveNext
    Wend
    If allsum < 10000 Then
        allsum = 0
        m_BillInterface.SetFieldValue "FInteger", 0
        m_BillInterface.SetFieldValue "FText1", "0"
    ElseIf allsum >= 10000 And allsum < 50000 Then
    m_BillInterface.SetFieldValue "FInteger", (allsum \ 10000)
        allsum = (allsum \ 10000) * 400
    m_BillInterface.SetFieldValue "FText1", "4%"
    ElseIf allsum >= 50000 And allsum < 100000 Then
    m_BillInterface.SetFieldValue "FInteger", (allsum \ 10000)
        allsum = (allsum \ 10000) * 500
    m_BillInterface.SetFieldValue "FText1", "5%"
    Else
    m_BillInterface.SetFieldValue "FInteger", (allsum \ 10000)
        allsum = (allsum \ 10000) * 600
    m_BillInterface.SetFieldValue "FText1", "6%"
    End If
    m_BillInterface.SetFieldValue "FDecimal6", allsum '累计奖励
    '触发 物料列值更新事件
    Dim blnResult As Boolean
    blnResult = m_BillInterface.BatchToDoActions("FBase1", -1, -1)
    
    MsgBox "读取完毕！", vbOKOnly, oK3.LoadKDString("金蝶提示")
    
    GoTo ExitSub
Catch:
    Set dct = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, oK3.LoadKDString("金蝶提示")
    End If
ExitSub:
    Set oK3 = Nothing
    Set rs = Nothing
    End If
    
End Sub


Private Sub m_BillInterface_Change(ByVal dct As KFO.IDictionary, ByVal dctFld As KFO.IDictionary, ByVal Col As Long, ByVal Row As Long, Cancel As Boolean)
 
    'TODO: 请在此处添加代码响应事件
    On Error GoTo Catch
    
    Dim strSQL As String
    Dim oK3  As K3ClassEvents.K3Lib
    Dim rs As ADODB.Recordset
    Set oK3 = m_BillInterface.K3Lib
    Dim bln As Boolean
    If UCase(dct.Value("FFieldName")) = UCase("FBase") Or UCase(dct.Value("FFieldName")) = UCase("FComboBox") Then
        bln = m_BillInterface.DeleteEntryData(2)
    ElseIf UCase(dct.Value("FFieldName")) = UCase("FBase1") Then
        m_BillInterface.SetFieldValue "FText2", "否", Row
    ElseIf UCase(dct.Value("FFieldName")) = UCase("FDecimal7") Then
        strSQL = " select a.FCoefficient from t_measureUnit a,t_icitem b where a.fitemid=b.fsaleunitid and b.fitemid= " & m_BillInterface.GetFieldValue("FBase1", Row, Enu_ValueType_FFLD)
        Set rs = oK3.GetData(strSQL)
        Dim FDecimal, FDecimal3, FDecimal5
        FDecimal = rs("FCoefficient") * dctFld("FFLD")
        FDecimal3 = FDecimal * IIf(m_BillInterface.GetFieldValue("FDecimal2", Row, Enu_ValueType_FFLD) = "", 0, m_BillInterface.GetFieldValue("FDecimal2", Row, Enu_ValueType_FFLD))
        FDecimal5 = IIf(m_BillInterface.GetFieldValue("FDecimal4", Row, Enu_ValueType_FFLD) = "", 0, m_BillInterface.GetFieldValue("FDecimal4", Row, Enu_ValueType_FFLD)) - FDecimal3
        m_BillInterface.SetFieldValue "FDecimal", FDecimal, Row
        m_BillInterface.SetFieldValue "FDecimal3", FDecimal3, Row
        m_BillInterface.SetFieldValue "FDecimal5", FDecimal5, Row
        m_BillInterface.SetFieldValue "FText2", "否", Row
        
        Set dct = m_BillInterface.TableInfo(m_BillInterface.TableInfo("Map")("FDecimal5")) 'm_BillInterface.GetFieldInfoByKey("FDecimal5")
        lSum = m_BillInterface.sum(dct, 1, m_BillInterface.BillEntrys(1).MaxRows)
        setmoney (lSum)
    ElseIf UCase(dct.Value("FFieldName")) = UCase("FDecimal8") Then
        strSQL = " select a.FCoefficient from t_measureUnit a,t_icitem b where a.fitemid=b.fsaleunitid and b.fitemid= " & m_BillInterface.GetFieldValue("FBase1", Row, Enu_ValueType_FFLD)
        Set rs = oK3.GetData(strSQL)
        Dim FDecimal1, FDecimal4
        FDecimal1 = rs("FCoefficient") * dctFld("FFLD")
        FDecimal4 = FDecimal1 * IIf(m_BillInterface.GetFieldValue("FDecimal2", Row, Enu_ValueType_FFLD) = "", 0, m_BillInterface.GetFieldValue("FDecimal2", Row, Enu_ValueType_FFLD))
        FDecimal5 = FDecimal4 - IIf(m_BillInterface.GetFieldValue("FDecimal3", Row, Enu_ValueType_FFLD) = "", 0, m_BillInterface.GetFieldValue("FDecimal3", Row, Enu_ValueType_FFLD))
        m_BillInterface.SetFieldValue "FDecimal1", FDecimal1, Row
        m_BillInterface.SetFieldValue "FDecimal4", FDecimal4, Row
        m_BillInterface.SetFieldValue "FDecimal5", FDecimal5, Row
        m_BillInterface.SetFieldValue "FText2", "否", Row
        Set dct = m_BillInterface.TableInfo(m_BillInterface.TableInfo("Map")("FDecimal5")) 'm_BillInterface.GetFieldInfoByKey("FDecimal5")
        lSum = m_BillInterface.sum(dct, 1, m_BillInterface.BillEntrys(1).MaxRows)
        setmoney (lSum)
    ElseIf UCase(dct.Value("FFieldName")) = UCase("FDecimal9") Then
        strSQL = " select a.FCoefficient from t_measureUnit a,t_icitem b where a.fitemid=b.fsaleunitid and b.fitemid= " & m_BillInterface.GetFieldValue("FBase1", Row, Enu_ValueType_FFLD)
        Set rs = oK3.GetData(strSQL)
        Dim FDecimal2
        FDecimal2 = dctFld("FFLD") / rs("FCoefficient")
        FDecimal3 = FDecimal2 * IIf(m_BillInterface.GetFieldValue("FDecimal", Row, Enu_ValueType_FFLD) = "", 0, m_BillInterface.GetFieldValue("FDecimal", Row, Enu_ValueType_FFLD))
        FDecimal4 = FDecimal2 * IIf(m_BillInterface.GetFieldValue("FDecimal1", Row, Enu_ValueType_FFLD) = "", 0, m_BillInterface.GetFieldValue("FDecimal1", Row, Enu_ValueType_FFLD))
        FDecimal5 = FDecimal4 - FDecimal3
        m_BillInterface.SetFieldValue "FDecimal2", FDecimal2, Row
        m_BillInterface.SetFieldValue "FDecimal3", FDecimal3, Row
        m_BillInterface.SetFieldValue "FDecimal4", FDecimal4, Row
        m_BillInterface.SetFieldValue "FDecimal5", FDecimal5, Row
        m_BillInterface.SetFieldValue "FText2", "否", Row
        Set dct = m_BillInterface.TableInfo(m_BillInterface.TableInfo("Map")("FDecimal5")) 'm_BillInterface.GetFieldInfoByKey("FDecimal5")
        lSum = m_BillInterface.sum(dct, 1, m_BillInterface.BillEntrys(1).MaxRows)
        setmoney (lSum)
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

Function setmoney(allsum As Double)
    If allsum < 10000 Then
        allsum = 0
        m_BillInterface.SetFieldValue "FInteger", 0
        m_BillInterface.SetFieldValue "FText1", "0"
    ElseIf allsum >= 10000 And allsum < 50000 Then
    m_BillInterface.SetFieldValue "FInteger", (allsum \ 10000)
        allsum = (allsum \ 10000) * 400
    m_BillInterface.SetFieldValue "FText1", "4%"
    ElseIf allsum >= 50000 And allsum < 100000 Then
    m_BillInterface.SetFieldValue "FInteger", (allsum \ 10000)
        allsum = (allsum \ 10000) * 500
    m_BillInterface.SetFieldValue "FText1", "5%"
    Else
    m_BillInterface.SetFieldValue "FInteger", (allsum \ 10000)
        allsum = (allsum \ 10000) * 600
    m_BillInterface.SetFieldValue "FText1", "6%"
    End If
    m_BillInterface.SetFieldValue "FDecimal6", allsum '累计奖励
End Function
