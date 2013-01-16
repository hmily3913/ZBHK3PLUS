VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Industry_LLJY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is Industry Bill Interface Class, made by K3BOSPLUGINSWIZAED"
 
'定义插件对象接口. 必须具有的声明, 以此来获得事件
Private WithEvents m_BillTransfer   As K3BillTransfer.Bill
Attribute m_BillTransfer.VB_VarHelpID = -1
 
Public Sub Show(ByVal oBillTransfer As Object)
 
    '接口实现
    '注意: 此方法必须存在, 请勿修改
    Set m_BillTransfer = oBillTransfer
 
End Sub

Private Sub Class_Terminate()
 
    '释放接口对象
    '注意: 此方法必须存在, 请勿修改
    Set m_BillTransfer = Nothing

End Sub
Private Sub m_BillTransfer_RetEvents(ByVal Para As KFO.IDictionary)
    On Error GoTo Err
    If Para("EventID") = "Bill_ButtonClick" Then
        If Para("Para")("Button").Key = "mnuUpSecCoef" And Para("EventIDProp") = 1 Then
            If m_BillTransfer.SaveVect.Item(1).Value("BillChecked") Then
    Dim objAPP As Object
    Dim Conn As ADODB.Connection
    Set Conn = New ADODB.Connection
    Conn.Open m_BillTransfer.Cnnstring
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    strSQL = "select a.*, b.fname from ICQCBill a inner join t_measureUnit b on a.FUnitID=b.fmeasureunitid and a.FCancellation=0 and a.FStatus=1 and a.FInterID=" & m_BillTransfer.SaveVect.Item(1).Value("FInterID")
    rs.Open strSQL, Conn, 1, 1
    If Not rs.EOF Then
        Dim FInStockInterID, FEntryID, fresult As Integer
        Dim fheadselft1257, fheadselft1258, fheadselft1259, FCheckQty As Double
        Dim FUnit, fheadselft1260 As String
        FInStockInterID = rs("FInStockInterID")
        FEntryID = rs("FSerialID")
        fresult = rs("fresult")
        FUnit = rs("fname")
        FCheckQty = rs("FCheckQty")
        fheadselft1260 = rs("fheadselft1260")
        fheadselft1257 = rs("fheadselft1257")
        fheadselft1258 = rs("fheadselft1258")
        fheadselft1259 = rs("fheadselft1259")
        If FUnit = "公斤(折百)" And fheadselft1259 <> fheadselft1257 And fheadselft1258 > 0 And fresult = 286 And fheadselft1260 <> "Y" And Round(FCheckQty, 4) = Round(fheadselft1257 * fheadselft1258, 4) Then
            Dim upQty As Double '差异数=（检测出来的比例-报检比例）*本次检验实物数量
            upQty = Math.Round((fheadselft1259 - fheadselft1257) * fheadselft1258, 4)
            '检验申请单数量检查
            Set rs = New ADODB.Recordset
            Dim FStockID, FItemID, FBatchNo As String
            strSQL = "select a.FCLOSED,b.* from POInStock a inner join POInStockEntry b on a.FInterID=b.FInterID and a.FInterID=" & FInStockInterID & " and b.FEntryID=" & FEntryID
            rs.Open strSQL, Conn, 1, 1
            If Not rs.EOF Then
                If rs("FCLOSED") = 1 Then
                    MsgBox "对应来料检验单已经关闭，更新失败！"
                    Exit Sub
                ElseIf rs("FQtyPass") + upQty < rs("FCommitQty") - rs("fconpassqty") + rs("FBackQty") Then
                    MsgBox "来料检验和个数<入库数量-让步接受数量+合格退料数量，更新失败！"
                    Exit Sub
                End If
                FStockID = rs("FStockID")
                FItemID = rs("FItemID")
                FBatchNo = rs("FBatchNo")
            End If
            '待检仓检查
            Set rs = New ADODB.Recordset
            strSQL = "select * from POInventory where  FStockID=" & FStockID & " and FBatchNo='" & FBatchNo & "' and FItemID=" & FItemID
            rs.Open strSQL, Conn, 1, 1
            If Not rs.EOF Then
                If rs("FQty") + upQty < 0 Or rs("FSecQty") = 0 Then
                    MsgBox "待检仓不允许为负数；辅助数量为0。更新失败！"
                    Exit Sub
                End If
            End If
            '开始执行更新事务
            Conn.BeginTrans
            '更新检验表
            Conn.Execute ("update ICQCBill set fheadselft1260='Y',FSendUpQty=FSendUpQty+" & upQty & ",FCheckQty=FCheckQty+" & upQty & ",FPassQty=FPassQty+" & upQty & ",FBasePassQty=FBasePassQty+" & upQty & " where FInterID=" & m_BillTransfer.SaveVect.Item(1).Value("FInterID"))
            '更新检验申请表
            strSQL = " UPDate t1 Set FAuxRelateQty = CASE WHEN Cast(isnull(FAuxRelateQty,0) as Float) + " & upQty & " < 0 THEN 0 ELSE Cast(isnull(FAuxRelateQty,0) as Float) + " & upQty & " END,FAuxQty = CASE WHEN Cast(isnull(FAuxQty,0) as Float) + " & upQty & " < 0 THEN 0 ELSE Cast(isnull(FAuxQty,0) as Float) + " & upQty & " END, FAuxQCheckQty = CASE WHEN Cast(isnull(FAuxQCheckQty,0) as Float) + " & upQty & " < 0 THEN 0 ELSE Cast(isnull(FAuxQCheckQty,0) as Float) + " & upQty & " END,  FAuxQtyPass = CASE WHEN Cast(isnull(FAuxQtyPass,0) as Float) + " & upQty & " < 0 THEN 0 ELSE Cast(isnull(FAuxQtyPass,0) as Float) + " & upQty & " END  From POInstockEntry t1  WHERE t1.FInterID = " & FInStockInterID & " AND t1.FEntryID = " & FEntryID & " AND t1.FItemID = " & FItemID & " "
            Conn.Execute (strSQL)
            strSQL = " Update t1 Set  t1.FQCheckQty = Cast(t1.FAuxQCheckQty as float) * Cast(t7.FCoefficient as Float), "
            strSQL = strSQL & "  t1.FQtyPass = Cast(t1.FAuxQtyPass as float) * Cast(t7.FCoefficient as Float),"
            strSQL = strSQL & "  t1.FQty = Cast(t1.FAuxQty as float) * Cast(t7.FCoefficient as Float),"
            strSQL = strSQL & "  t1.Famount = Cast(t1.FAuxQty as float) * Cast(t7.FCoefficient as Float)*FPrice,"
            strSQL = strSQL & "  t1.FRelateQty = Cast(t1.FAuxRelateQty as float) * Cast(t7.FCoefficient as Float)"
            strSQL = strSQL & "  FROM  POInstockEntry t1"
            strSQL = strSQL & "  INNER JOIN t_measureUnit t7 on t1.FUnitID = t7.FItemID"
            strSQL = strSQL & "  INNER JOIN t_icitem t4 on t1.FItemID = t4.FItemID"
            strSQL = strSQL & "  WHERE t1.FInterID = " & FInStockInterID & " AND t1.FEntryID = " & FEntryID & " AND t1.FItemID = " & FItemID & " "
            Conn.Execute (strSQL)
            '更新虚仓即时库存
            strSQL = "update POInventory set FQty=FQty+" & upQty & " where  FStockID=" & FStockID & " and FBatchNo='" & FBatchNo & "' and FItemID=" & FItemID
            Conn.Execute (strSQL)
            '更新采购订单反写
            strSQL = "update pn SET FCommitQty=ISNULL(pn.FCommitQty,0)+" & upQty & ",FAuxCommitQty=ISNULL(pn.FAuxCommitQty,0)+" & upQty & "/ISNULL(m.FCoefficient,1) FROM poorderEntry pn inner join POInStockEntry t on t.FSourceInterID=pn.FInterID AND t.FSourceEntryID=pn.FEntryID and t.FSourceTranType=71 and t.FInterID = " & FInStockInterID & " AND t.FEntryID = " & FEntryID & " AND t.FItemID = " & FItemID & " LEFT JOIN t_MeasureUnit m ON m.FItemID=pn.FUnitID "
            Conn.Execute (strSQL)
            '更新采购发票（没什么用）
            strSQL = "update pn SET FCommitQty=ISNULL(pn.FCommitQty,0)+" & upQty & ",FAuxCommitQty=ISNULL(pn.FAuxCommitQty,0)+" & upQty & "/ISNULL(m.FCoefficient,1) FROM ICPurchaseEntry pn inner join POInStockEntry t on t.FSourceInterID=pn.FInterID AND t.FSourceEntryID=pn.FEntryID and t.FSourceTranType=71 and t.FInterID = " & FInStockInterID & " AND t.FEntryID = " & FEntryID & " AND t.FItemID = " & FItemID & " LEFT JOIN t_MeasureUnit m ON m.FItemID=pn.FUnitID "
            Conn.Execute (strSQL)
            
            '   询问用户是否想提交以上所做的全部更改。
            If MsgBox("确定提交更新？ ", vbYesNo) = vbYes Then
                  Conn.CommitTrans
                  MsgBox "更新成功！"
            Else
                  Conn.RollbackTrans
            End If
        Else
            MsgBox "不允许更新！原因：单位不是折百；检验比例与报检比例相同；检验数量为0；检验不合格；已经更新过；检验数量（折百）与实物数量*报检比例不相同；"
        End If
    End If
    Set Conn = Nothing
    Set objAPP = Nothing
    
            Else
                MsgBox "当前单据未审核，不允许进行比例更新！"
            End If
        End If
    End If
    Exit Sub
Err:
    MsgBox "更新数据库失败，请联系IT人员"
End Sub

Private Sub m_BillTransfer_LoadBillEnd(ByVal ShowType As Long)
'    If m_BillTransfer.SaveVect.Item(1).Value("BillChecked") Then
        '增加按钮
        With m_BillTransfer.BillForm.tlbTool.Buttons
            .Add .Count + 1, "mnuUpSecCoef", "更新比例", , 21
        End With

'    End If
    '新增时
'    Dim objAPP As Object
'    Dim adoRS As ADODB.Recordset
'    Dim strSql As String

    
'    Dim strFInStockInterID As String
'    Dim strFCheckQty As String
'    Dim strFEntryID As String
'    Dim FUnitID As String

'    Dim iFInStockInterID As Integer
'    Dim iFZDW As Integer
'    Dim iFZSL As Integer
'    Dim iFZHSL As Integer
'    Dim iJCBL As Integer
'    Dim iCHECKQTY As Integer
    
    On Error GoTo Err
    
'    Set objAPP = CreateObject("K3MAppconnection.AppConnection")
    
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'    Dim cpHeadCtl As Variant
'    cpHeadCtl = m_BillTransfer.HeadCtl
'    For i = LBound(m_BillTransfer.HeadCtl) To UBound(m_BillTransfer.HeadCtl)
'         If m_BillTransfer.HeadCtl(i).FieldName = "FInStockInterID" Then
'             strFInStockInterID = m_BillTransfer.HeadCtl(i).InterID
'             iFInStockInterID = i
'         End If
''         If m_BillTransfer.HeadCtl(i).FieldName = "FBizType" Then
''              strFBizType = m_BillTransfer.HeadCtl(i).FinterID
''         End If
'         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1256" Then
'            m_BillTransfer.Head(i).Locked = True
'            m_BillTransfer.Head(i).Enabled = False
'             iFZDW = i
'         End If
'         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1257" Then
'             iFZHSL = i
'         End If
'         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1258" Then
'             iFZSL = i
'         End If
'         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1259" Then
'             iJCBL = i
'         End If
'         If m_BillTransfer.HeadCtl(i).FieldName = "FSerialID" Then
'             strFEntryID = m_BillTransfer.HeadCtl(i).vale
'         End If
'         If m_BillTransfer.HeadCtl(i).FieldName = "FCheckQty" Then
'             strFCheckQty = m_BillTransfer.HeadCtl(i).vale
'             iCHECKQTY = i
'         End If
'         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "funitid" Then
'            FUnitID = m_BillTransfer.HeadCtl(i).Name
'         End If
         
'    Next i
'    m_BillTransfer.HeadCtl = cpHeadCtl
    '读取检验申请单中的对应辅助计量单位信息
    '新增时
'If ShowType = 0 And FUnitID = "公斤(折百)" Then
'    If strFInStockInterID <> "" And strFEntryID <> "" Then
'        strSql = "select b.FSecCoefficient,b.FSecQty,b.FQty,b.fsecbackqty from POInStockEntry AS b inner join POInStock AS a on b.FInterID = a.FInterID and a.ftrantype=702 where a.finterid=" & strFInStockInterID & " and b.FEntryID=" & strFEntryID & " and b.FRelateQty<b.FQty"
'        Set adoRS = objAPP.getanyrecordset(m_BillTransfer.Cnnstring, strSql)
'        If Not adoRS.EOF Then
'            If adoRS("FSecCoefficient") <> 0 Then
'              Dim checkqty As Double
'                checkqty = adoRS("FSecQty") - adoRS("fsecbackqty")
'                If adoRS("FQty") <> CDbl(strFCheckQty) Then checkqty = CDbl(strFCheckQty) / adoRS("FSecCoefficient")
'                m_BillTransfer.SetHead iFZHSL, adoRS("FSecCoefficient")
'                m_BillTransfer.SetHead iJCBL, adoRS("FSecCoefficient")
'                m_BillTransfer.SetHead iFZSL, Math.Round(checkqty, 4)
'                m_BillTransfer.Head(iCHECKQTY).Locked = True
'                m_BillTransfer.Head(iCHECKQTY).Enabled = False
'            End If
'        End If
'    End If
'Else
If FUnitID = "公斤(折百)" And m_BillTransfer.HeadCtl(iFZHSL).vale <> "0.0000" And m_BillTransfer.HeadCtl(iFZHSL).vale <> "" And Not m_BillTransfer.SaveVect.Item(1).Value("BillChecked") Then
'    m_BillTransfer.SetHead iCHECKQTY, Math.Round(CDbl(m_BillTransfer.HeadCtl(iFZSL).vale) * CDbl(m_BillTransfer.HeadCtl(iFZHSL).vale), 4)
    m_BillTransfer.Head(iCHECKQTY).Locked = True
    m_BillTransfer.Head(iCHECKQTY).Enabled = False
End If
'    Set adoRS = Nothing
'    Set objAPP = Nothing
    Exit Sub
Err:
    MsgBox "更新数据库失败，请联系IT人员"
End Sub
'Private Sub m_BillTransfer_BeforeSave(ByVal bNew As Boolean, ReturnCode As Long)
'    Dim objAPP As Object
'    Dim adoRS As ADODB.Recordset
'    Dim strSql As String'

    
'    Dim strFInStockInterID As String
'    Dim strFCheckQty As String
'    Dim strFEntryID As String
'    Dim FUnitID As String
    
'    Dim iFInStockInterID As Integer
'    Dim iFZDW As Integer
'    Dim iFZSL As Integer
'    Dim iFZHSL As Integer
    
    
'    On Error GoTo Err
    
'    Set objAPP = CreateObject("K3MAppconnection.AppConnection")
    
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'    For i = LBound(m_BillTransfer.HeadCtl) To UBound(m_BillTransfer.HeadCtl)
'         If m_BillTransfer.HeadCtl(i).FieldName = "FInStockInterID" Then
'             strFInStockInterID = m_BillTransfer.HeadCtl(i).InterID
'             iFInStockInterID = i
'         End If
''         If m_BillTransfer.HeadCtl(i).FieldName = "FBizType" Then
''              strFBizType = m_BillTransfer.HeadCtl(i).FinterID
''         End If
'         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1256" Then
'             iFZDW = i
'         End If
'         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1257" Then
'             iFZHSL = i
'         End If
'         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1258" Then
'             iFZSL = i
'         End If
'         If m_BillTransfer.HeadCtl(i).FieldName = "FSerialID" Then
'             strFEntryID = m_BillTransfer.HeadCtl(i).vale
'         End If
'         If m_BillTransfer.HeadCtl(i).FieldName = "FCheckQty" Then
'             strFCheckQty = m_BillTransfer.HeadCtl(i).vale
'         End If
'         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "funitid" Then
'            FUnitID = m_BillTransfer.HeadCtl(i).Name
'         End If
'    Next i
    
'    '读取检验申请单中的对应辅助计量单位信息
'    If strFInStockInterID <> "" Then
'        strSql = "select b.FSecCoefficient,b.FSecQty,b.FQty,b.fsecbackqty from POInStockEntry AS b inner join POInStock AS a on b.FInterID = a.FInterID and a.ftrantype=702 where a.finterid=" & strFInStockInterID & " and b.FEntryID=" & strFEntryID & " "
'        Set adoRS = objAPP.getanyrecordset(m_BillTransfer.Cnnstring, strSql)
'        If Not adoRS.EOF Then
'            If adoRS("FSecCoefficient") <> 0 And FUnitID = "公斤(折百)" Then
'                Dim checkqty As Double
'                checkqty = adoRS("FSecQty") - adoRS("fsecbackqty")
'                If adoRS("FQty") <> CDbl(strFCheckQty) Then checkqty = CDbl(strFCheckQty) / adoRS("FSecCoefficient")
'                m_BillTransfer.SetHead iFZHSL, adoRS("FSecCoefficient")
'                m_BillTransfer.SetHead iFZSL, Math.Round(checkqty, 4)
'            End If
'        End If
'    End If
    
'    Set adoRS = Nothing
'    Set objAPP = Nothing
'    Exit Sub
'Err:
'    MsgBox "更新数据库失败，请联系IT人员"

'End Sub
Private Sub m_BillTransfer_HeadChange(ByVal CtlIndex As Long, ByVal Value As Variant, ByVal bNewBill As Boolean, Cancel As Boolean)
    Dim objAPP As Object
    Dim adoRS As ADODB.Recordset
    Dim strSQL As String

    
    Dim strFInStockInterID As String
    Dim strFCheckQty As String
    Dim strFEntryID As String
    Dim FUnitID As String

    Dim iFInStockInterID As Integer
    Dim iFZDW As Integer
    Dim iFZSL As Integer
    Dim iFZHSL As Integer
    Dim iJCBL As Integer
    Dim iCHECKQTY, iFPassQty As Integer
    Dim fheadselft1257 As String
    
    If LCase(m_BillTransfer.HeadCtl(CtlIndex).FieldName) = "fheadselft1258" Then
        For i = LBound(m_BillTransfer.HeadCtl) To UBound(m_BillTransfer.HeadCtl)
             If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1257" Then
                fheadselft1257 = m_BillTransfer.HeadCtl(i).vale
                 iFZHSL = i
             End If
             If m_BillTransfer.HeadCtl(i).FieldName = "FCheckQty" Then
                 iCHECKQTY = i
             End If
         If m_BillTransfer.HeadCtl(i).FieldName = "FPassQty" Then
             iFPassQty = i
         End If
            If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "funitid" Then
               FUnitID = m_BillTransfer.HeadCtl(i).Name
            End If
        Next i
        If fheadselft1257 <> "" And fheadselft1257 <> "0.0000" And FUnitID = "公斤(折百)" Then
            m_BillTransfer.SetHead iCHECKQTY, Math.Round(CDbl(Value) * CDbl(fheadselft1257), 4)
            m_BillTransfer.SetHead iFPassQty, Math.Round(CDbl(Value) * CDbl(fheadselft1257), 4)
        End If
    ElseIf m_BillTransfer.HeadCtl(CtlIndex).FieldName = "FInStockInterID" Then
    
    On Error GoTo Err
    
    Set objAPP = CreateObject("K3MAppconnection.AppConnection")
    
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'    Dim cpHeadCtl As Variant
'    cpHeadCtl = m_BillTransfer.HeadCtl
    For i = LBound(m_BillTransfer.HeadCtl) To UBound(m_BillTransfer.HeadCtl)
         If m_BillTransfer.HeadCtl(i).FieldName = "FInStockInterID" Then
             strFInStockInterID = m_BillTransfer.HeadCtl(i).InterID
             iFInStockInterID = i
         End If
'         If m_BillTransfer.HeadCtl(i).FieldName = "FBizType" Then
'              strFBizType = m_BillTransfer.HeadCtl(i).FinterID
'         End If
         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1256" Then
            m_BillTransfer.Head(i).Locked = True
            m_BillTransfer.Head(i).Enabled = False
             iFZDW = i
         End If
         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1257" Then
             iFZHSL = i
         End If
         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1258" Then
             iFZSL = i
         End If
         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1259" Then
             iJCBL = i
         End If
         If m_BillTransfer.HeadCtl(i).FieldName = "FSerialID" Then
             strFEntryID = m_BillTransfer.HeadCtl(i).vale
         End If
         If m_BillTransfer.HeadCtl(i).FieldName = "FCheckQty" Then
             strFCheckQty = m_BillTransfer.HeadCtl(i).vale
             iCHECKQTY = i
         End If
         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "funitid" Then
            FUnitID = m_BillTransfer.HeadCtl(i).Name
         End If
         
    Next i
'    m_BillTransfer.HeadCtl = cpHeadCtl
    '读取检验申请单中的对应辅助计量单位信息
    '新增时
If ShowType = 0 And FUnitID = "公斤(折百)" Then
    If strFInStockInterID <> "" And strFEntryID <> "" Then
        strSQL = "select b.FSecCoefficient,b.FSecQty,b.FQty,b.fsecbackqty from POInStockEntry AS b inner join POInStock AS a on b.FInterID = a.FInterID and a.ftrantype=702 where a.finterid=" & strFInStockInterID & " and b.FEntryID=" & strFEntryID & " and b.FRelateQty<b.FQty"
        Set adoRS = objAPP.getanyrecordset(m_BillTransfer.Cnnstring, strSQL)
        If Not adoRS.EOF Then
            If adoRS("FSecCoefficient") <> 0 Then
              Dim checkqty As Double
                checkqty = adoRS("FSecQty") - adoRS("fsecbackqty")
                If adoRS("FQty") <> CDbl(strFCheckQty) Then checkqty = CDbl(strFCheckQty) / adoRS("FSecCoefficient")
                m_BillTransfer.SetHead iFZHSL, adoRS("FSecCoefficient")
                m_BillTransfer.SetHead iJCBL, adoRS("FSecCoefficient")
                m_BillTransfer.SetHead iFZSL, Math.Round(checkqty, 4)
                m_BillTransfer.Head(iCHECKQTY).Locked = True
                m_BillTransfer.Head(iCHECKQTY).Enabled = False
            End If
        End If
    End If
ElseIf FUnitID = "公斤(折百)" And m_BillTransfer.HeadCtl(iFZHSL).vale <> "0.0000" And m_BillTransfer.HeadCtl(iFZHSL).vale <> "" And Not m_BillTransfer.SaveVect.Item(1).Value("BillChecked") Then
    m_BillTransfer.SetHead iCHECKQTY, Math.Round(CDbl(m_BillTransfer.HeadCtl(iFZSL).vale) * CDbl(m_BillTransfer.HeadCtl(iFZHSL).vale), 4)
    m_BillTransfer.Head(iCHECKQTY).Locked = True
    m_BillTransfer.Head(iCHECKQTY).Enabled = False
End If
    Set adoRS = Nothing
    Set objAPP = Nothing
    Exit Sub
Err:
    MsgBox "更新数据库失败，请联系IT人员"
    Exit Sub
    
    End If
End Sub
Private Sub m_BillTransfer_RefreshControl()
    Dim FUnitID, fheadselft1257 As String
    Dim iCHECKQTY As Integer
    For i = LBound(m_BillTransfer.HeadCtl) To UBound(m_BillTransfer.HeadCtl)
         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1257" Then
             iFZHSL = i
             fheadselft1257 = m_BillTransfer.HeadCtl(i).vale
         End If
         If m_BillTransfer.HeadCtl(i).FieldName = "FCheckQty" Then
             iCHECKQTY = i
         End If
         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "funitid" Then
            FUnitID = m_BillTransfer.HeadCtl(i).Name
         End If
    Next i
    If FUnitID = "公斤(折百)" And fheadselft1257 <> "" And fheadselft1257 <> "0.0000" Then 'CDbl(fheadselft1257) > 0 And
            m_BillTransfer.Head(iCHECKQTY).Locked = True
            m_BillTransfer.Head(iCHECKQTY).Enabled = False
    End If
End Sub
Private Sub m_BillTransfer_BeforeSelBills(ByVal Para As KFO.IDictionary)
    Dim objAPP As Object
    Dim adoRS As ADODB.Recordset
    Dim strSQL As String

    
    Dim strFInStockInterID As String
    Dim strFCheckQty As String
    Dim strFEntryID As String
    Dim FUnitID As String

    Dim iFInStockInterID As Integer
    Dim iFZDW As Integer
    Dim iFZSL As Integer
    Dim iFZHSL As Integer
    Dim iJCBL As Integer
    Dim iCHECKQTY, iFPassQty As Integer
    Dim fheadselft1257 As String
    
    On Error GoTo Err
    
    Set objAPP = CreateObject("K3MAppconnection.AppConnection")
    
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'    Dim cpHeadCtl As Variant
'    cpHeadCtl = m_BillTransfer.HeadCtl
    strFEntryID = Para("rsData").Fields("FEntryID")
    strFInStockInterID = Para("rsData").Fields("FInterID")
    FUnitID = Para("rsData").Fields("FUnitName")
    For i = LBound(m_BillTransfer.HeadCtl) To UBound(m_BillTransfer.HeadCtl)
         If m_BillTransfer.HeadCtl(i).FieldName = "FInStockInterID" Then
             iFInStockInterID = i
         End If
         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1256" Then
            m_BillTransfer.Head(i).Locked = True
            m_BillTransfer.Head(i).Enabled = False
             iFZDW = i
         End If
         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1257" Then
             iFZHSL = i
         End If
         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1258" Then
             iFZSL = i
         End If
         If LCase(m_BillTransfer.HeadCtl(i).FieldName) = "fheadselft1259" Then
             iJCBL = i
         End If
         If m_BillTransfer.HeadCtl(i).FieldName = "FCheckQty" Then
             strFCheckQty = m_BillTransfer.HeadCtl(i).vale
             iCHECKQTY = i
         End If
         
    Next i
'    m_BillTransfer.HeadCtl = cpHeadCtl
    '读取检验申请单中的对应辅助计量单位信息
    '新增时
If FUnitID = "公斤(折百)" Then
    If strFInStockInterID <> "" And strFEntryID <> "" Then
        strSQL = "select b.FSecCoefficient,b.FSecQty,b.FQty,b.fsecbackqty from POInStockEntry AS b inner join POInStock AS a on b.FInterID = a.FInterID and a.ftrantype=702 where a.finterid=" & strFInStockInterID & " and b.FEntryID=" & strFEntryID & " and b.FRelateQty<b.FQty"
        Set adoRS = objAPP.getanyrecordset(m_BillTransfer.Cnnstring, strSQL)
        If Not adoRS.EOF Then
            If adoRS("FSecCoefficient") <> 0 Then
              Dim checkqty As Double
                checkqty = adoRS("FSecQty") - adoRS("fsecbackqty")
'                If adoRS("FQty") <> CDbl(strFCheckQty) Then checkqty = CDbl(strFCheckQty) / adoRS("FSecCoefficient")
                m_BillTransfer.SetHead iFZHSL, adoRS("FSecCoefficient")
                m_BillTransfer.SetHead iJCBL, adoRS("FSecCoefficient")
                m_BillTransfer.SetHead iFZSL, Math.Round(checkqty, 4)
                m_BillTransfer.Head(iCHECKQTY).Locked = True
                m_BillTransfer.Head(iCHECKQTY).Enabled = False
            End If
        End If
    End If
End If
    Set adoRS = Nothing
    Set objAPP = Nothing
    Exit Sub
Err:
    MsgBox "更新数据库失败，请联系IT人员"
 


End Sub
