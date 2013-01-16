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
 
'����������ӿ�. ������е�����, �Դ�������¼�
Private WithEvents m_BillTransfer   As K3BillTransfer.Bill
Attribute m_BillTransfer.VB_VarHelpID = -1
 
Public Sub Show(ByVal oBillTransfer As Object)
 
    '�ӿ�ʵ��
    'ע��: �˷����������, �����޸�
    Set m_BillTransfer = oBillTransfer
 
End Sub

Private Sub Class_Terminate()
 
    '�ͷŽӿڶ���
    'ע��: �˷����������, �����޸�
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
        If FUnit = "����(�۰�)" And fheadselft1259 <> fheadselft1257 And fheadselft1258 > 0 And fresult = 286 And fheadselft1260 <> "Y" And Round(FCheckQty, 4) = Round(fheadselft1257 * fheadselft1258, 4) Then
            Dim upQty As Double '������=���������ı���-���������*���μ���ʵ������
            upQty = Math.Round((fheadselft1259 - fheadselft1257) * fheadselft1258, 4)
            '�������뵥�������
            Set rs = New ADODB.Recordset
            Dim FStockID, FItemID, FBatchNo As String
            strSQL = "select a.FCLOSED,b.* from POInStock a inner join POInStockEntry b on a.FInterID=b.FInterID and a.FInterID=" & FInStockInterID & " and b.FEntryID=" & FEntryID
            rs.Open strSQL, Conn, 1, 1
            If Not rs.EOF Then
                If rs("FCLOSED") = 1 Then
                    MsgBox "��Ӧ���ϼ��鵥�Ѿ��رգ�����ʧ�ܣ�"
                    Exit Sub
                ElseIf rs("FQtyPass") + upQty < rs("FCommitQty") - rs("fconpassqty") + rs("FBackQty") Then
                    MsgBox "���ϼ���͸���<�������-�ò���������+�ϸ���������������ʧ�ܣ�"
                    Exit Sub
                End If
                FStockID = rs("FStockID")
                FItemID = rs("FItemID")
                FBatchNo = rs("FBatchNo")
            End If
            '����ּ��
            Set rs = New ADODB.Recordset
            strSQL = "select * from POInventory where  FStockID=" & FStockID & " and FBatchNo='" & FBatchNo & "' and FItemID=" & FItemID
            rs.Open strSQL, Conn, 1, 1
            If Not rs.EOF Then
                If rs("FQty") + upQty < 0 Or rs("FSecQty") = 0 Then
                    MsgBox "����ֲ�����Ϊ��������������Ϊ0������ʧ�ܣ�"
                    Exit Sub
                End If
            End If
            '��ʼִ�и�������
            Conn.BeginTrans
            '���¼����
            Conn.Execute ("update ICQCBill set fheadselft1260='Y',FSendUpQty=FSendUpQty+" & upQty & ",FCheckQty=FCheckQty+" & upQty & ",FPassQty=FPassQty+" & upQty & ",FBasePassQty=FBasePassQty+" & upQty & " where FInterID=" & m_BillTransfer.SaveVect.Item(1).Value("FInterID"))
            '���¼��������
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
            '������ּ�ʱ���
            strSQL = "update POInventory set FQty=FQty+" & upQty & " where  FStockID=" & FStockID & " and FBatchNo='" & FBatchNo & "' and FItemID=" & FItemID
            Conn.Execute (strSQL)
            '���²ɹ�������д
            strSQL = "update pn SET FCommitQty=ISNULL(pn.FCommitQty,0)+" & upQty & ",FAuxCommitQty=ISNULL(pn.FAuxCommitQty,0)+" & upQty & "/ISNULL(m.FCoefficient,1) FROM poorderEntry pn inner join POInStockEntry t on t.FSourceInterID=pn.FInterID AND t.FSourceEntryID=pn.FEntryID and t.FSourceTranType=71 and t.FInterID = " & FInStockInterID & " AND t.FEntryID = " & FEntryID & " AND t.FItemID = " & FItemID & " LEFT JOIN t_MeasureUnit m ON m.FItemID=pn.FUnitID "
            Conn.Execute (strSQL)
            '���²ɹ���Ʊ��ûʲô�ã�
            strSQL = "update pn SET FCommitQty=ISNULL(pn.FCommitQty,0)+" & upQty & ",FAuxCommitQty=ISNULL(pn.FAuxCommitQty,0)+" & upQty & "/ISNULL(m.FCoefficient,1) FROM ICPurchaseEntry pn inner join POInStockEntry t on t.FSourceInterID=pn.FInterID AND t.FSourceEntryID=pn.FEntryID and t.FSourceTranType=71 and t.FInterID = " & FInStockInterID & " AND t.FEntryID = " & FEntryID & " AND t.FItemID = " & FItemID & " LEFT JOIN t_MeasureUnit m ON m.FItemID=pn.FUnitID "
            Conn.Execute (strSQL)
            
            '   ѯ���û��Ƿ����ύ����������ȫ�����ġ�
            If MsgBox("ȷ���ύ���£� ", vbYesNo) = vbYes Then
                  Conn.CommitTrans
                  MsgBox "���³ɹ���"
            Else
                  Conn.RollbackTrans
            End If
        Else
            MsgBox "���������£�ԭ�򣺵�λ�����۰٣���������뱨�������ͬ����������Ϊ0�����鲻�ϸ��Ѿ����¹��������������۰٣���ʵ������*�����������ͬ��"
        End If
    End If
    Set Conn = Nothing
    Set objAPP = Nothing
    
            Else
                MsgBox "��ǰ����δ��ˣ����������б������£�"
            End If
        End If
    End If
    Exit Sub
Err:
    MsgBox "�������ݿ�ʧ�ܣ�����ϵIT��Ա"
End Sub

Private Sub m_BillTransfer_LoadBillEnd(ByVal ShowType As Long)
'    If m_BillTransfer.SaveVect.Item(1).Value("BillChecked") Then
        '���Ӱ�ť
        With m_BillTransfer.BillForm.tlbTool.Buttons
            .Add .Count + 1, "mnuUpSecCoef", "���±���", , 21
        End With

'    End If
    '����ʱ
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
    '��ȡ�������뵥�еĶ�Ӧ����������λ��Ϣ
    '����ʱ
'If ShowType = 0 And FUnitID = "����(�۰�)" Then
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
If FUnitID = "����(�۰�)" And m_BillTransfer.HeadCtl(iFZHSL).vale <> "0.0000" And m_BillTransfer.HeadCtl(iFZHSL).vale <> "" And Not m_BillTransfer.SaveVect.Item(1).Value("BillChecked") Then
'    m_BillTransfer.SetHead iCHECKQTY, Math.Round(CDbl(m_BillTransfer.HeadCtl(iFZSL).vale) * CDbl(m_BillTransfer.HeadCtl(iFZHSL).vale), 4)
    m_BillTransfer.Head(iCHECKQTY).Locked = True
    m_BillTransfer.Head(iCHECKQTY).Enabled = False
End If
'    Set adoRS = Nothing
'    Set objAPP = Nothing
    Exit Sub
Err:
    MsgBox "�������ݿ�ʧ�ܣ�����ϵIT��Ա"
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
    
'    '��ȡ�������뵥�еĶ�Ӧ����������λ��Ϣ
'    If strFInStockInterID <> "" Then
'        strSql = "select b.FSecCoefficient,b.FSecQty,b.FQty,b.fsecbackqty from POInStockEntry AS b inner join POInStock AS a on b.FInterID = a.FInterID and a.ftrantype=702 where a.finterid=" & strFInStockInterID & " and b.FEntryID=" & strFEntryID & " "
'        Set adoRS = objAPP.getanyrecordset(m_BillTransfer.Cnnstring, strSql)
'        If Not adoRS.EOF Then
'            If adoRS("FSecCoefficient") <> 0 And FUnitID = "����(�۰�)" Then
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
'    MsgBox "�������ݿ�ʧ�ܣ�����ϵIT��Ա"

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
        If fheadselft1257 <> "" And fheadselft1257 <> "0.0000" And FUnitID = "����(�۰�)" Then
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
    '��ȡ�������뵥�еĶ�Ӧ����������λ��Ϣ
    '����ʱ
If ShowType = 0 And FUnitID = "����(�۰�)" Then
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
ElseIf FUnitID = "����(�۰�)" And m_BillTransfer.HeadCtl(iFZHSL).vale <> "0.0000" And m_BillTransfer.HeadCtl(iFZHSL).vale <> "" And Not m_BillTransfer.SaveVect.Item(1).Value("BillChecked") Then
    m_BillTransfer.SetHead iCHECKQTY, Math.Round(CDbl(m_BillTransfer.HeadCtl(iFZSL).vale) * CDbl(m_BillTransfer.HeadCtl(iFZHSL).vale), 4)
    m_BillTransfer.Head(iCHECKQTY).Locked = True
    m_BillTransfer.Head(iCHECKQTY).Enabled = False
End If
    Set adoRS = Nothing
    Set objAPP = Nothing
    Exit Sub
Err:
    MsgBox "�������ݿ�ʧ�ܣ�����ϵIT��Ա"
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
    If FUnitID = "����(�۰�)" And fheadselft1257 <> "" And fheadselft1257 <> "0.0000" Then 'CDbl(fheadselft1257) > 0 And
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
    '��ȡ�������뵥�еĶ�Ӧ����������λ��Ϣ
    '����ʱ
If FUnitID = "����(�۰�)" Then
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
    MsgBox "�������ݿ�ʧ�ܣ�����ϵIT��Ա"
 


End Sub