VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BillEvent_KHZKJSD"
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
    
    If m_BillInterface.GetFieldValue("FDate", , Enu_ValueType_FFLD) = "" Or m_BillInterface.GetFieldValue("FDate1", , Enu_ValueType_FFLD) = "" Or m_BillInterface.GetFieldValue("FBase", , Enu_ValueType_FFLD) = "" Then
        MsgBox "开始日期、截至日期、业务员不能为空，读取失败！", vbCritical, oK3.LoadKDString("金蝶提示")
        Exit Sub
    End If
    If m_BillInterface.GetFieldValue("FUser", , Enu_ValueType_FFLD) <> "" Then
        MsgBox "审核状态不允许进行此操作！", vbCritical, oK3.LoadKDString("金蝶提示")
        Exit Sub
    End If
'    strSQL = "select 1 from z_KHZKJS where (('" & m_BillInterface.GetFieldValue("FDate", , Enu_ValueType_FFLD) & "'<=FDate1 and '" & m_BillInterface.GetFieldValue("FDate", , Enu_ValueType_FFLD) & "'>=FDate) or ('" & m_BillInterface.GetFieldValue("FDate1", , Enu_ValueType_FFLD) & "'>=Fdate and '" & m_BillInterface.GetFieldValue("FDate1", , Enu_ValueType_FFLD) & "'<=Fdate1)) and FBase=" & m_BillInterface.GetFieldValue("FBase", , Enu_ValueType_FFLD) & " and FID<>" & m_BillInterface.CurBillID
'    Set rs = oK3.GetData(strSQL)
'    If Not rs.EOF Then
'        MsgBox "该业务员在此日期区间内已经做过结算单，不允许重复生成！", vbCritical, oK3.LoadKDString("金蝶提示")
'        Exit Sub
'    End If
    '删除现在分录
'    Dim bln As Boolean
'    bln = m_BillInterface.DeleteEntryData(2)
    
    strSQL = "select a.*,isnull(b.FVALUE,0)*a.FauxQTY as dicprice,a.ZK-isnull(b.FVALUE,0)*a.FauxQTY as dismoney,isnull(c.FauxQTY,0) as qtsl from ( "
    strSQL = strSQL & " select t2.fitemid,t5.FNumber,t4.FNumber as FCustID,t3.FPRICE,SUM(case when round(t2.Fauxtaxprice,4)>=t3.FPRICE then T2.FQTY else 0 end) AS FQTY, "
    strSQL = strSQL & " SUM(case when round(t2.Fauxtaxprice,4)>=t3.FPRICE then T2.FauxQTY else 0 end) AS FauxQTY,"
    strSQL = strSQL & " SUM(case when round(t2.Fauxtaxprice,4)>=t3.FPRICE then ISNULL(T2.FAmountincludetax,0)+ISNULL(T2.FAMTDISCOUNT,0) else 0 end) AS SK,"
    strSQL = strSQL & " SUM(case when round(t2.Fauxtaxprice,4)>=t3.FPRICE then ISNULL(T2.FAmountincludetax,0)+ISNULL(T2.FENTRYSELFI0463,0) else 0 end) AS JS,SUM(ISNULL(T2.FAMTDISCOUNT,0)-ISNULL(T2.FENTRYSELFI0463,0)) AS ZK"
    strSQL = strSQL & " from ICSALE T1,ICSALEENTRY T2,t_Organization t4,t_icitem t5,"
    strSQL = strSQL & " (SELECT TOP 1 FITEMID,FPRICE,FBEGDATE,FENDDATE FROM ICPRCPLYENTRY where FChecked=1 AND FBEGDATE<='" & m_BillInterface.GetFieldValue("FDate", , Enu_ValueType_FFLD) & "' AND FItemID= " & m_BillInterface.GetFieldValue("FBase4", , Enu_ValueType_FFLD) & " order by FBEGDATE desc) T3"
    strSQL = strSQL & " Where T1.FINTERID = T2.FINTERID"
    strSQL = strSQL & " AND T1.FTRANTYPE=80 AND T1.FCheckerID<>0 AND T1.FCANCELLATION=0"
    strSQL = strSQL & " AND T2.FITEMID=T3.FITEMID and t2.FItemid=t5.Fitemid and t1.FCustID=t4.fitemid"
    strSQL = strSQL & " AND T1.FDATE>='" & m_BillInterface.GetFieldValue("FDate", , Enu_ValueType_FFLD) & "' AND T1.FDATE<='" & m_BillInterface.GetFieldValue("FDate1", , Enu_ValueType_FFLD) & "'"
    strSQL = strSQL & "  AND T2.FItemID= " & m_BillInterface.GetFieldValue("FBase4", , Enu_ValueType_FFLD) & " "
    strSQL = strSQL & "  AND T4.femployee= " & m_BillInterface.GetFieldValue("FBase", , Enu_ValueType_FFLD) & " "
'    strSQL = strSQL & "  AND T1.FCustID= " & m_BillInterface.GetFieldValue("FBase3", Row, Enu_ValueType_FFLD) & " "
    strSQL = strSQL & " group by t2.fitemid,t5.FNumber,t4.FNumber,t3.FPRICE"
    strSQL = strSQL & " HAVING SUM(ISNULL(T2.FAMTDISCOUNT,0)-ISNULL(T2.FENTRYSELFI0463,0))<>0 ) a"
    strSQL = strSQL & " left join (select t2.fitemid,t5.FNumber,t4.FNumber as FCustID,t3.FPRICE,SUM(case when round(t2.Fauxtaxprice,4)>=t3.FPRICE then T2.FQTY else 0 end) AS FQTY, "
    strSQL = strSQL & " SUM(case when round(t2.Fauxtaxprice,4)>=t3.FPRICE then T2.FauxQTY else 0 end) AS FauxQTY,"
    strSQL = strSQL & " SUM(case when round(t2.Fauxtaxprice,4)>=t3.FPRICE then ISNULL(T2.FAmountincludetax,0)+ISNULL(T2.FAMTDISCOUNT,0) else 0 end) AS SK,"
    strSQL = strSQL & " SUM(case when round(t2.Fauxtaxprice,4)>=t3.FPRICE then ISNULL(T2.FAmountincludetax,0)+ISNULL(T2.FENTRYSELFI0463,0) else 0 end) AS JS,SUM(ISNULL(T2.FAMTDISCOUNT,0)-ISNULL(T2.FENTRYSELFI0463,0)) AS ZK"
    strSQL = strSQL & " from ICSALE T1,ICSALEENTRY T2,t_Organization t4,t_icitem t5,t_subsItem t6,"
    strSQL = strSQL & " (select FITEMID,FPRICE,Row_Number() over (partition by FITEMID order by FBEGDATE desc) as rn from ICPRCPLYENTRY where FChecked=1 AND FBEGDATE<='" & m_BillInterface.GetFieldValue("FDate", , Enu_ValueType_FFLD) & "') T3"
    strSQL = strSQL & " Where T1.FINTERID = T2.FINTERID and t3.rn=1 and (t2.FITEMID=t6.fitemid or T2.FITEMID=T6.FSUBsItemID) "
    strSQL = strSQL & " AND T1.FTRANTYPE=80 AND T1.FCheckerID<>0 AND T1.FCANCELLATION=0  AND T2.FItemID<> " & m_BillInterface.GetFieldValue("FBase4", , Enu_ValueType_FFLD) & " "
    strSQL = strSQL & " AND T2.FITEMID=T3.FITEMID and t2.FItemid=t5.Fitemid and t1.FCustID=t4.fitemid"
    strSQL = strSQL & " AND T1.FDATE>='" & m_BillInterface.GetFieldValue("FDate", , Enu_ValueType_FFLD) & "' AND T1.FDATE<='" & m_BillInterface.GetFieldValue("FDate1", , Enu_ValueType_FFLD) & "'"
    strSQL = strSQL & "  AND (T6.FSUBsItemID= " & m_BillInterface.GetFieldValue("FBase4", , Enu_ValueType_FFLD) & " or T6.FItemID= " & m_BillInterface.GetFieldValue("FBase4", , Enu_ValueType_FFLD) & ") "
    strSQL = strSQL & "  AND T4.femployee= " & m_BillInterface.GetFieldValue("FBase", , Enu_ValueType_FFLD) & " "
'    strSQL = strSQL & "  AND T1.FCustID= " & m_BillInterface.GetFieldValue("FBase3", Row, Enu_ValueType_FFLD) & " "
    strSQL = strSQL & " group by t2.fitemid,t5.FNumber,t4.FNumber,t3.FPRICE"
    strSQL = strSQL & " HAVING SUM(ISNULL(T2.FAMTDISCOUNT,0)-ISNULL(T2.FENTRYSELFI0463,0))<>0 ) c on a.FCustID=c.FCustID"
    strSQL = strSQL & " left join ICDisPlyEntry b on a.fitemid=b.fitemid and isnull(a.FauxQTY,0)+isnull(c.FauxQTY,0)>=b.FBEGQTY"
    strSQL = strSQL & " and isnull(a.FauxQTY,0)+isnull(c.FauxQTY,0)<=FENDQTY and b.FBEGDATE<='" & m_BillInterface.GetFieldValue("FDate", , Enu_ValueType_FFLD) & "' "
    Set rs = oK3.GetData(strSQL)
    
    Dim i As Long
    i = 1
    While (Not rs.EOF)
        '批量插入行数据
        If rs("FAUXQTY") = 0 Then
            m_BillInterface.InsertNewRowAndFill 2, i, "FBase1", rs("FNUMBER"), "FBase3", rs("fcustid"), "FDecimal", rs("FAUXQTY"), "FDecimal1", rs("SK"), "FDecimal2", rs("SK"), "FDecimal3", rs("JS"), "FDecimal4", rs("JS"), "FDecimal5", rs("ZK"), "FDecimal6", rs("FQTY"), "FDecimal7", rs("dicprice"), "FDecimal8", rs("dismoney")
        Else
            m_BillInterface.InsertNewRowAndFill 2, i, "FBase1", rs("FNUMBER"), "FBase3", rs("fcustid"), "FDecimal", rs("FAUXQTY"), "FDecimal1", rs("SK") / rs("FAUXQTY"), "FDecimal2", rs("SK"), "FDecimal3", rs("JS") / rs("FAUXQTY"), "FDecimal4", rs("JS"), "FDecimal5", rs("ZK"), "FDecimal6", rs("FQTY"), "FDecimal7", rs("dicprice"), "FDecimal8", rs("dismoney")
        End If
        i = i + 1
        rs.MoveNext
    Wend
    
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
    If Col = -1 And Row = -1 Then
        Dim bln As Boolean
        If dctFld.Count > 3 Then
            If Left(dctFld("FFND"), 1) <> "P" Then
                bln = m_BillInterface.DeleteEntryData(2)
            End If
        Else
            bln = m_BillInterface.DeleteEntryData(2)
        End If
    ElseIf Col = 4 Or Col = 7 Then
        If m_BillInterface.GetFieldValue("FDate", , Enu_ValueType_FFLD) = "" Or m_BillInterface.GetFieldValue("FDate1", , Enu_ValueType_FFLD) = "" Then
            MsgBox "开始日期、截至日期不能为空，读取失败！", vbCritical, "金蝶提示"
            Exit Sub
        End If
        If m_BillInterface.GetFieldValue("FBase1", , Enu_ValueType_FFLD) <> "" And m_BillInterface.GetFieldValue("FBase3", , Enu_ValueType_FFLD) <> "" Then
            On Error GoTo Catch
            Dim strSQL As String
            Dim rs As ADODB.Recordset
            Set oK3 = m_BillInterface.K3Lib
            
'    strSQL = "select 1 from z_KHZKJS a,z_KHZKJSEntry b where a.fid=b.fid and (('" & m_BillInterface.GetFieldValue("FDate", , Enu_ValueType_FFLD) & "'<=FDate1 and '" & m_BillInterface.GetFieldValue("FDate", , Enu_ValueType_FFLD) & "'>=FDate) or ('" & m_BillInterface.GetFieldValue("FDate1", , Enu_ValueType_FFLD) & "'>=Fdate and '" & m_BillInterface.GetFieldValue("FDate1", , Enu_ValueType_FFLD) & "'<=Fdate1)) and FBase1=" & m_BillInterface.GetFieldValue("FBase1", Row, Enu_ValueType_FFLD) & " and FBase3=" & m_BillInterface.GetFieldValue("FBase3", Row, Enu_ValueType_FFLD) & " and a.FID<>" & m_BillInterface.FID
'    Set rs = oK3.GetData(strSQL)
'    If Not rs.EOF Then
'        MsgBox "相同客户、产品在此日期区间内已经做过结算单，不允许重复生成！", vbCritical, oK3.LoadKDString("金蝶提示")
'        GoTo ExitSub
'    End If
    strSQL = "select a.*,isnull(b.FVALUE,0)*FauxQTY as dicprice,ZK-isnull(b.FVALUE,0)*FauxQTY as dismoney from ( "
    strSQL = strSQL & " select t2.fitemid,t1.FCustID,t3.FPRICE,SUM(case when round(t2.Fauxtaxprice,4)>=t3.FPRICE then T2.FQTY else 0 end) AS FQTY, "
    strSQL = strSQL & " SUM(case when round(t2.Fauxtaxprice,4)>=t3.FPRICE then T2.FauxQTY else 0 end) AS FauxQTY,"
    strSQL = strSQL & " SUM(case when round(t2.Fauxtaxprice,4)>=t3.FPRICE then ISNULL(T2.FAmountincludetax,0)+ISNULL(T2.FAMTDISCOUNT,0) else 0 end) AS SK,"
    strSQL = strSQL & " SUM(case when round(t2.Fauxtaxprice,4)>=t3.FPRICE then ISNULL(T2.FAmountincludetax,0)+ISNULL(T2.FENTRYSELFI0463,0) else 0 end) AS JS,SUM(ISNULL(T2.FAMTDISCOUNT,0)-ISNULL(T2.FENTRYSELFI0463,0)) AS ZK"
    strSQL = strSQL & " from ICSALE T1,ICSALEENTRY T2,"
    strSQL = strSQL & " (SELECT TOP 1 FITEMID,FPRICE,FBEGDATE,FENDDATE FROM ICPRCPLYENTRY where FChecked=1 AND FBEGDATE<='" & m_BillInterface.GetFieldValue("FDate", , Enu_ValueType_FFLD) & "' AND FItemID= " & m_BillInterface.GetFieldValue("FBase1", Row, Enu_ValueType_FFLD) & ") T3"
    strSQL = strSQL & " Where T1.FINTERID = T2.FINTERID"
    strSQL = strSQL & " AND T1.FTRANTYPE=80 AND T1.FCheckerID<>0 AND T1.FCANCELLATION=0"
    strSQL = strSQL & " AND T2.FITEMID=T3.FITEMID"
    strSQL = strSQL & " AND T1.FDATE>='" & m_BillInterface.GetFieldValue("FDate", , Enu_ValueType_FFLD) & "' AND T1.FDATE<='" & m_BillInterface.GetFieldValue("FDate1", , Enu_ValueType_FFLD) & "'"
    strSQL = strSQL & "  AND T2.FItemID= " & m_BillInterface.GetFieldValue("FBase1", Row, Enu_ValueType_FFLD) & " "
    strSQL = strSQL & "  AND T1.FCustID= " & m_BillInterface.GetFieldValue("FBase3", Row, Enu_ValueType_FFLD) & " "
    strSQL = strSQL & " group by t2.fitemid,t1.FCustID,t3.FPRICE"
    strSQL = strSQL & " HAVING SUM(ISNULL(T2.FAMTDISCOUNT,0)-ISNULL(T2.FENTRYSELFI0463,0))<>0 ) a"
    strSQL = strSQL & " left join ICDisPlyEntry b on a.fitemid=b.fitemid and a.FauxQTY>=b.FBEGQTY"
    strSQL = strSQL & " and a.FauxQTY<=FENDQTY and b.FBEGDATE<='" & m_BillInterface.GetFieldValue("FDate", , Enu_ValueType_FFLD) & "' "
    Set rs = oK3.GetData(strSQL)
    If rs.EOF Then
        MsgBox "对应客户、产品在此日期区间内无可结算的折扣，请检测！", vbCritical, oK3.LoadKDString("金蝶提示")
        GoTo ExitSub
    Else
        m_BillInterface.SetFieldValue "FDecimal", rs("FAUXQTY"), Row
        m_BillInterface.SetFieldValue "FDecimal2", rs("SK"), Row
        m_BillInterface.SetFieldValue "FDecimal4", rs("JS"), Row
        m_BillInterface.SetFieldValue "FDecimal5", rs("ZK"), Row
        m_BillInterface.SetFieldValue "FDecimal6", rs("FQTY"), Row
        m_BillInterface.SetFieldValue "FDecimal7", rs("dicprice"), Row
        m_BillInterface.SetFieldValue "FDecimal8", rs("dismoney"), Row
        If rs("FAUXQTY") = 0 Then
            m_BillInterface.SetFieldValue "FDecimal1", rs("SK"), Row
            m_BillInterface.SetFieldValue "FDecimal3", rs("JS"), Row
        Else
            m_BillInterface.SetFieldValue "FDecimal1", rs("SK") / rs("FAUXQTY"), Row
            m_BillInterface.SetFieldValue "FDecimal3", rs("JS") / rs("FAUXQTY"), Row
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
        End If
    End If
End Sub

