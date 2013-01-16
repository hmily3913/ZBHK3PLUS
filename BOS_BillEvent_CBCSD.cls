VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BillEvent_CBCSD"
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
    
    If m_BillInterface.GetFieldValue("FComboBox", , Enu_ValueType_FFLD) = "" Or m_BillInterface.GetFieldValue("FComboBox1", , Enu_ValueType_FFLD) = "" Or m_BillInterface.GetFieldValue("FComboBox2", , Enu_ValueType_FFLD) = "" Or m_BillInterface.GetFieldValue("FBase", , Enu_ValueType_FFLD) = "" Then
        MsgBox "年份、月份、物料不能为空，计算失败！", vbCritical, oK3.LoadKDString("金蝶提示")
        Exit Sub
    End If
    If m_BillInterface.GetFieldValue("FUser", , Enu_ValueType_FFLD) <> "" Then
        MsgBox "审核状态不允许计算，计算失败！", vbCritical, oK3.LoadKDString("金蝶提示")
        Exit Sub
    End If
    '删除现在分录
    Dim bln As Boolean
    bln = m_BillInterface.DeleteEntryData(2)
    
    strSQL = "Select t1.FNumber,a.FQty,case when t7.FBase is not null then isnull(t7.FDecimal5,0) else isnull(t9.FDecimal3,0) end as ccf, "
    strSQL = strSQL & " case when t7.FBase is not null then isnull(t7.FDecimal,0) else isnull(t8.直接人工,0) end AS 直接人工,"
    strSQL = strSQL & " case when t7.FBase is not null then isnull(t7.FDecimal1,0) else isnull(t8.电费,0) end AS 电费,"
    strSQL = strSQL & " case when t7.FBase is not null then isnull(t7.FDecimal13,0) else isnull(t8.折旧费,0) end AS 折旧费,"
    strSQL = strSQL & " case when t7.FBase is not null then isnull(t7.FDecimal12,0) else isnull(t8.激光刻码,0) end AS 激光刻码,"
    strSQL = strSQL & " case when t7.FBase is not null then isnull(t7.FDecimal3,0) else isnull(t8.每吨制造费分摊,0) end AS 每吨制造费分摊,"
    strSQL = strSQL & " t1.FTaxRate,t6.fnumber as punit,a.fbomnumber "
    strSQL = strSQL & " From icbom a inner join t_ICItem t1 on t1.FItemID = a.fitemid and a.FInterID=t1.FCostBOMID and t1.FItemid=" & m_BillInterface.GetFieldValue("FBase", , Enu_ValueType_FFLD)
    strSQL = strSQL & " left join t_item t4 on t1.fnumber=t4.fnumber and t4.fitemclassid='2001' "
    strSQL = strSQL & " left join t_MeasureUnit t6 on t6.fMeasureUnitid=t1.fsaleunitid "
    strSQL = strSQL & " left join (select t2.*,Row_Number() over (partition by t2.FBase order by t2.FDate desc) as rn "
    strSQL = strSQL & " from z_JHJGCB t1, z_JHJGCBEntry t2 where t1.FID=t2.FID and t1.FUser>0 AND year(T2.FDATE)=" & m_BillInterface.GetFieldValue("FComboBox2", , Enu_ValueType_FFLD) & ") t7 on t7.FBase=t1.FItemid and t7.rn=1 "
    strSQL = strSQL & " left join (select *,Row_Number() over (partition by 产品代码 order by 年份 desc) as rn "
    strSQL = strSQL & " from z_YCJGCB where 年份= '" & m_BillInterface.GetFieldValue("FComboBox2", , Enu_ValueType_FFLD) & "' ) t8 on t8.产品代码=t1.Fnumber and t8.rn=1 "
    strSQL = strSQL & "  left join z_CCFBZ t9 on t9.fuser>0 and FComboBox='" & m_BillInterface.GetFieldValue("FComboBox2", , Enu_ValueType_FFLD) & "' "
    Set rs = oK3.GetData(strSQL)
    Dim df, sgtbf, cjzzfy, zjf, ccf, qty, xiaoji As Double
    Dim fnumber As String
    Dim bzjgf, taxrate, xj, xsdj1, xsdj2, zjrg2 As Double
    ccf = rs("ccf"): qty = rs("FQty")
    xiaoji = Round(rs("电费"), 2) + Round(rs("激光刻码"), 2) + Round(rs("折旧费"), 2) + Round(rs("直接人工"), 2) + Round(rs("每吨制造费分摊"), 2)
    fnumber = rs("FNumber")
    zjrg2 = Round(rs("直接人工"), 2) + Round(rs("激光刻码"), 2)
    bzjgf = rs("每吨制造费分摊"): taxrate = rs("FTaxRate")
    
    m_BillInterface.SetFieldValue "FDecimal8", rs("激光刻码")
    m_BillInterface.SetFieldValue "FDecimal10", rs("折旧费")
    m_BillInterface.SetFieldValue "FDecimal4", rs("直接人工") '直接人工
    m_BillInterface.SetFieldValue "FDecimal12", rs("电费") '直接人工
    m_BillInterface.SetFieldValue "FDecimal28", ccf '仓储费
    
    
    m_BillInterface.InsertNewRowAndFill 2, 1, "FText", ".0", "FBase1", fnumber, "FBase2", rs("punit"), "FQty", 1000, "FPrice", 0, "FAmount", 0, "FDecimal22", 0, "FText2", "", "FText1", rs("fbomnumber"), "FDecimal24", 0, "FDecimal26", 0, "FDecimal27", rs("直接人工"), "FDecimal29", rs("每吨制造费分摊"), "FDecimal30", rs("折旧费"), "FDecimal31", rs("电费")
    
    strSQL = "exec [Z_成本计算] '" & fnumber & "','" & m_BillInterface.GetFieldValue("FComboBox", , Enu_ValueType_FFLD) & "','" & m_BillInterface.GetFieldValue("FComboBox1", , Enu_ValueType_FFLD) & "','" & m_BillInterface.GetFieldValue("FComboBox2", , Enu_ValueType_FFLD) & "' "
    Set rs = oK3.GetData(strSQL)
    
    Dim i As Long
    Dim TotalCB, BZWCB, YCLCB, BCPJGF, jxse, zjrg As Double
    TotalCB = 0: BZWCB = 0: YCLCB = 0: BCPJGF = 0: jxse = 0: zjrg = 0
    i = 2
    While (Not rs.EOF)
        '批量插入行数据
        Dim t_jgcb, t_zjrg
        t_jgcb = 0
        t_zjrg = 0
        If rs("BomNumber") = "" Then
            TotalCB = TotalCB + rs("FAmount") / qty * 1000 / (1 + rs("taxrate") / 100)
            jxse = jxse + rs("FAmount") / qty * 1000 / (1 + rs("taxrate") / 100) * rs("taxrate") / 100
            If rs("BZWFlag") = "是" Then
                BZWCB = BZWCB + rs("FAmount") / qty * 1000 / (1 + rs("taxrate") / 100)
            Else
                YCLCB = YCLCB + rs("FAmount") / qty * 1000 / (1 + rs("taxrate") / 100)
            End If
            m_BillInterface.InsertNewRowAndFill 2, i, "FText", rs("FLevel"), "FBase1", rs("FNumber"), "FBase2", rs("FUnitNMB"), "FQty", rs("FQty") / qty * 1000, "FPrice", rs("FPrice"), "FAmount", rs("FAmount") / qty * 1000, "FDecimal22", rs("taxrate"), "FText2", rs("BZWFlag"), "FText1", rs("BomNumber"), "FDecimal24", rs("FAmount") / qty * 10 / (1 + rs("taxrate") / 100) * rs("taxrate"), "FDecimal26", rs("FAmount") / qty * 1000 / (1 + rs("taxrate") / 100), "FDecimal27", 0
        Else
            t_jgcb = rs("每吨制造费分摊") / qty * rs("FQty")
            t_zjrg = rs("直接人工") / qty * rs("FQty")
            BCPJGF = BCPJGF + Round(t_jgcb, 2)
            zjrg = zjrg + Round(t_zjrg, 2)
            m_BillInterface.InsertNewRowAndFill 2, i, "FText", rs("FLevel"), "FBase1", rs("FNumber"), "FBase2", rs("FUnitNMB"), "FQty", rs("FQty") / qty * 1000, "FPrice", 0, "FAmount", 0, "FDecimal22", 0, "FText2", rs("BZWFlag"), "FText1", rs("BomNumber"), "FDecimal24", 0, "FDecimal26", 0, "FDecimal27", t_zjrg, "FDecimal29", t_jgcb, "FDecimal30", rs("折旧费"), "FDecimal31", rs("电费")
        End If
        i = i + 1
        rs.MoveNext
    Wend
    
    '触发 物料列值更新事件
    Dim blnResult As Boolean
    blnResult = m_BillInterface.BatchToDoActions("FBase1", -1, -1)
    xiaoji = xiaoji + zjrg + BCPJGF
    '填写单头
    m_BillInterface.SetFieldValue "FDecimal", YCLCB '原辅料成本
    m_BillInterface.SetFieldValue "FDecimal1", BZWCB '包装物成本
    m_BillInterface.SetFieldValue "FDecimal2", Round((BZWCB + YCLCB) * 0.005, 2) '材料损耗
    m_BillInterface.SetFieldValue "FDecimal3", zjrg '半成品直接人工
    m_BillInterface.SetFieldValue "FDecimal11", (BZWCB + YCLCB) * 1.005 '直接材料成本合计
    m_BillInterface.SetFieldValue "FDecimal9", BCPJGF + bzjgf '总的制造费用
    m_BillInterface.SetFieldValue "FDecimal5", zjrg + zjrg2 '加工费合计
    m_BillInterface.SetFieldValue "FDecimal13", xiaoji '生产成本合计
    m_BillInterface.SetFieldValue "FDecimal6", Round(((YCLCB + BZWCB) * 1.005 + xiaoji + ccf) * 0.021, 2)  '资金成本
    m_BillInterface.SetFieldValue "FDecimal7", Round(((YCLCB + BZWCB) * 1.005 + xiaoji + ccf) * 1.021, 2)  '成本合计
    m_BillInterface.SetFieldValue "FDecimal21", jxse '进项税额
    m_BillInterface.SetFieldValue "FDecimal23", qty '换算比例
    m_BillInterface.DoAction "FDecimal14"
    m_BillInterface.DoAction "FCheckBox"
    '文件价格
    Dim z01, z02, z03, z04, z05 As Double
'    If xsdj1 > 0 Then
'        z01 = xsdj1 / (1 + taxrate / 100)
'        z02 = z01 * taxrate / 100 - jxse
'        z03 = z02 * 0.12 + z01 * (0.001 + 0.8 * 0.0003)
'        z04 = z01 - ((YCLCB + BZWCB) * 1.005 + xiaoji + BCPJGF + ccf) * 1.021 - z03
'        z05 = z04 / z01 * 100
'        m_BillInterface.SetFieldValue "FDecimal8", xsdj1 '销售价格含税
'        m_BillInterface.SetFieldValue "FDecimal9", Round(z01, 2) '销售价格不含税
'        m_BillInterface.SetFieldValue "FDecimal10", z02 '应交增值税
'        m_BillInterface.SetFieldValue "FDecimal11", Round(z03, 2) '附加税
'        m_BillInterface.SetFieldValue "FDecimal12", Round(z04, 2) '毛利额
'        m_BillInterface.SetFieldValue "FDecimal13", Round(z05, 2) '毛利率
'        m_BillInterface.SetFieldValue "FDecimal25", xsdj1 / 1000 * qty '箱价
'    End If
    '预测价格
'    If xsdj2 > 0 Then
'        z01 = xsdj2 / (1 + taxrate / 100)
'        z02 = z01 * taxrate / 100 - jxse
'        z03 = z02 * 0.12 + z01 * (0.001 + 0.8 * 0.0003)
'        z04 = z01 - ((YCLCB + BZWCB) * 1.005 + xiaoji + BCPJGF + ccf) * 1.021 - z03
'        z05 = z04 / z01 * 100
'        m_BillInterface.SetFieldValue "FDecimal14", xsdj2 '销售价格含税
'        m_BillInterface.SetFieldValue "FDecimal15", Round(z01, 2) '销售价格不含税
'        m_BillInterface.SetFieldValue "FDecimal16", z02 '销售价格含税
'        m_BillInterface.SetFieldValue "FDecimal17", Round(z03, 2) '附加税
'        m_BillInterface.SetFieldValue "FDecimal18", Round(z04, 2) '毛利额
'        m_BillInterface.SetFieldValue "FDecimal19", Round(z05, 2) '毛利率
'        m_BillInterface.SetFieldValue "FDecimal20", xj '箱价
'    End If
    MsgBox "计算完毕！", vbOKOnly, oK3.LoadKDString("金蝶提示")
    
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
