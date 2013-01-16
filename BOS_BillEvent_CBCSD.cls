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
 
'���� BillEvent �ӿ�. ������е�����, �Դ�������¼�
Private WithEvents m_BillInterface  As BillEvent
Attribute m_BillInterface.VB_VarHelpID = -1
 
Public Sub Show(ByVal oBillInterface As Object)
 
    'BillEvent �ӿ�ʵ��
    'ע��: �˷����������, �����޸�
    Set m_BillInterface = oBillInterface
 
End Sub

Private Sub Class_Terminate()
 
    '�ͷŽӿڶ���
    'ע��: �˷����������, �����޸�
    Set m_BillInterface = Nothing

End Sub

Private Sub m_BillInterface_CommandClick(ByVal CommandName As String, bCancel As Boolean)
 
    'TODO: ���ڴ˴����Ӵ�����Ӧ�¼�
    '��Ӧ���㰴ť
    If CommandName = "FButton" Then
    On Error GoTo Catch
    
    Dim strSQL As String
    Dim strSQLBatch() As String
'    Dim i As Integer
    Dim oK3  As K3ClassEvents.K3Lib
    Dim rs As ADODB.Recordset
    Set oK3 = m_BillInterface.K3Lib
    
    If m_BillInterface.GetFieldValue("FComboBox", , Enu_ValueType_FFLD) = "" Or m_BillInterface.GetFieldValue("FComboBox1", , Enu_ValueType_FFLD) = "" Or m_BillInterface.GetFieldValue("FComboBox2", , Enu_ValueType_FFLD) = "" Or m_BillInterface.GetFieldValue("FBase", , Enu_ValueType_FFLD) = "" Then
        MsgBox "��ݡ��·ݡ����ϲ���Ϊ�գ�����ʧ�ܣ�", vbCritical, oK3.LoadKDString("�����ʾ")
        Exit Sub
    End If
    If m_BillInterface.GetFieldValue("FUser", , Enu_ValueType_FFLD) <> "" Then
        MsgBox "���״̬���������㣬����ʧ�ܣ�", vbCritical, oK3.LoadKDString("�����ʾ")
        Exit Sub
    End If
    'ɾ�����ڷ�¼
    Dim bln As Boolean
    bln = m_BillInterface.DeleteEntryData(2)
    
    strSQL = "Select t1.FNumber,a.FQty,case when t7.FBase is not null then isnull(t7.FDecimal5,0) else isnull(t9.FDecimal3,0) end as ccf, "
    strSQL = strSQL & " case when t7.FBase is not null then isnull(t7.FDecimal,0) else isnull(t8.ֱ���˹�,0) end AS ֱ���˹�,"
    strSQL = strSQL & " case when t7.FBase is not null then isnull(t7.FDecimal1,0) else isnull(t8.���,0) end AS ���,"
    strSQL = strSQL & " case when t7.FBase is not null then isnull(t7.FDecimal13,0) else isnull(t8.�۾ɷ�,0) end AS �۾ɷ�,"
    strSQL = strSQL & " case when t7.FBase is not null then isnull(t7.FDecimal12,0) else isnull(t8.�������,0) end AS �������,"
    strSQL = strSQL & " case when t7.FBase is not null then isnull(t7.FDecimal3,0) else isnull(t8.ÿ������ѷ�̯,0) end AS ÿ������ѷ�̯,"
    strSQL = strSQL & " t1.FTaxRate,t6.fnumber as punit,a.fbomnumber "
    strSQL = strSQL & " From icbom a inner join t_ICItem t1 on t1.FItemID = a.fitemid and a.FInterID=t1.FCostBOMID and t1.FItemid=" & m_BillInterface.GetFieldValue("FBase", , Enu_ValueType_FFLD)
    strSQL = strSQL & " left join t_item t4 on t1.fnumber=t4.fnumber and t4.fitemclassid='2001' "
    strSQL = strSQL & " left join t_MeasureUnit t6 on t6.fMeasureUnitid=t1.fsaleunitid "
    strSQL = strSQL & " left join (select t2.*,Row_Number() over (partition by t2.FBase order by t2.FDate desc) as rn "
    strSQL = strSQL & " from z_JHJGCB t1, z_JHJGCBEntry t2 where t1.FID=t2.FID and t1.FUser>0 AND year(T2.FDATE)=" & m_BillInterface.GetFieldValue("FComboBox2", , Enu_ValueType_FFLD) & ") t7 on t7.FBase=t1.FItemid and t7.rn=1 "
    strSQL = strSQL & " left join (select *,Row_Number() over (partition by ��Ʒ���� order by ��� desc) as rn "
    strSQL = strSQL & " from z_YCJGCB where ���= '" & m_BillInterface.GetFieldValue("FComboBox2", , Enu_ValueType_FFLD) & "' ) t8 on t8.��Ʒ����=t1.Fnumber and t8.rn=1 "
    strSQL = strSQL & "  left join z_CCFBZ t9 on t9.fuser>0 and FComboBox='" & m_BillInterface.GetFieldValue("FComboBox2", , Enu_ValueType_FFLD) & "' "
    Set rs = oK3.GetData(strSQL)
    Dim df, sgtbf, cjzzfy, zjf, ccf, qty, xiaoji As Double
    Dim fnumber As String
    Dim bzjgf, taxrate, xj, xsdj1, xsdj2, zjrg2 As Double
    ccf = rs("ccf"): qty = rs("FQty")
    xiaoji = Round(rs("���"), 2) + Round(rs("�������"), 2) + Round(rs("�۾ɷ�"), 2) + Round(rs("ֱ���˹�"), 2) + Round(rs("ÿ������ѷ�̯"), 2)
    fnumber = rs("FNumber")
    zjrg2 = Round(rs("ֱ���˹�"), 2) + Round(rs("�������"), 2)
    bzjgf = rs("ÿ������ѷ�̯"): taxrate = rs("FTaxRate")
    
    m_BillInterface.SetFieldValue "FDecimal8", rs("�������")
    m_BillInterface.SetFieldValue "FDecimal10", rs("�۾ɷ�")
    m_BillInterface.SetFieldValue "FDecimal4", rs("ֱ���˹�") 'ֱ���˹�
    m_BillInterface.SetFieldValue "FDecimal12", rs("���") 'ֱ���˹�
    m_BillInterface.SetFieldValue "FDecimal28", ccf '�ִ���
    
    
    m_BillInterface.InsertNewRowAndFill 2, 1, "FText", ".0", "FBase1", fnumber, "FBase2", rs("punit"), "FQty", 1000, "FPrice", 0, "FAmount", 0, "FDecimal22", 0, "FText2", "", "FText1", rs("fbomnumber"), "FDecimal24", 0, "FDecimal26", 0, "FDecimal27", rs("ֱ���˹�"), "FDecimal29", rs("ÿ������ѷ�̯"), "FDecimal30", rs("�۾ɷ�"), "FDecimal31", rs("���")
    
    strSQL = "exec [Z_�ɱ�����] '" & fnumber & "','" & m_BillInterface.GetFieldValue("FComboBox", , Enu_ValueType_FFLD) & "','" & m_BillInterface.GetFieldValue("FComboBox1", , Enu_ValueType_FFLD) & "','" & m_BillInterface.GetFieldValue("FComboBox2", , Enu_ValueType_FFLD) & "' "
    Set rs = oK3.GetData(strSQL)
    
    Dim i As Long
    Dim TotalCB, BZWCB, YCLCB, BCPJGF, jxse, zjrg As Double
    TotalCB = 0: BZWCB = 0: YCLCB = 0: BCPJGF = 0: jxse = 0: zjrg = 0
    i = 2
    While (Not rs.EOF)
        '��������������
        Dim t_jgcb, t_zjrg
        t_jgcb = 0
        t_zjrg = 0
        If rs("BomNumber") = "" Then
            TotalCB = TotalCB + rs("FAmount") / qty * 1000 / (1 + rs("taxrate") / 100)
            jxse = jxse + rs("FAmount") / qty * 1000 / (1 + rs("taxrate") / 100) * rs("taxrate") / 100
            If rs("BZWFlag") = "��" Then
                BZWCB = BZWCB + rs("FAmount") / qty * 1000 / (1 + rs("taxrate") / 100)
            Else
                YCLCB = YCLCB + rs("FAmount") / qty * 1000 / (1 + rs("taxrate") / 100)
            End If
            m_BillInterface.InsertNewRowAndFill 2, i, "FText", rs("FLevel"), "FBase1", rs("FNumber"), "FBase2", rs("FUnitNMB"), "FQty", rs("FQty") / qty * 1000, "FPrice", rs("FPrice"), "FAmount", rs("FAmount") / qty * 1000, "FDecimal22", rs("taxrate"), "FText2", rs("BZWFlag"), "FText1", rs("BomNumber"), "FDecimal24", rs("FAmount") / qty * 10 / (1 + rs("taxrate") / 100) * rs("taxrate"), "FDecimal26", rs("FAmount") / qty * 1000 / (1 + rs("taxrate") / 100), "FDecimal27", 0
        Else
            t_jgcb = rs("ÿ������ѷ�̯") / qty * rs("FQty")
            t_zjrg = rs("ֱ���˹�") / qty * rs("FQty")
            BCPJGF = BCPJGF + Round(t_jgcb, 2)
            zjrg = zjrg + Round(t_zjrg, 2)
            m_BillInterface.InsertNewRowAndFill 2, i, "FText", rs("FLevel"), "FBase1", rs("FNumber"), "FBase2", rs("FUnitNMB"), "FQty", rs("FQty") / qty * 1000, "FPrice", 0, "FAmount", 0, "FDecimal22", 0, "FText2", rs("BZWFlag"), "FText1", rs("BomNumber"), "FDecimal24", 0, "FDecimal26", 0, "FDecimal27", t_zjrg, "FDecimal29", t_jgcb, "FDecimal30", rs("�۾ɷ�"), "FDecimal31", rs("���")
        End If
        i = i + 1
        rs.MoveNext
    Wend
    
    '���� ������ֵ�����¼�
    Dim blnResult As Boolean
    blnResult = m_BillInterface.BatchToDoActions("FBase1", -1, -1)
    xiaoji = xiaoji + zjrg + BCPJGF
    '��д��ͷ
    m_BillInterface.SetFieldValue "FDecimal", YCLCB 'ԭ���ϳɱ�
    m_BillInterface.SetFieldValue "FDecimal1", BZWCB '��װ��ɱ�
    m_BillInterface.SetFieldValue "FDecimal2", Round((BZWCB + YCLCB) * 0.005, 2) '�������
    m_BillInterface.SetFieldValue "FDecimal3", zjrg '���Ʒֱ���˹�
    m_BillInterface.SetFieldValue "FDecimal11", (BZWCB + YCLCB) * 1.005 'ֱ�Ӳ��ϳɱ��ϼ�
    m_BillInterface.SetFieldValue "FDecimal9", BCPJGF + bzjgf '�ܵ��������
    m_BillInterface.SetFieldValue "FDecimal5", zjrg + zjrg2 '�ӹ��Ѻϼ�
    m_BillInterface.SetFieldValue "FDecimal13", xiaoji '�����ɱ��ϼ�
    m_BillInterface.SetFieldValue "FDecimal6", Round(((YCLCB + BZWCB) * 1.005 + xiaoji + ccf) * 0.021, 2)  '�ʽ�ɱ�
    m_BillInterface.SetFieldValue "FDecimal7", Round(((YCLCB + BZWCB) * 1.005 + xiaoji + ccf) * 1.021, 2)  '�ɱ��ϼ�
    m_BillInterface.SetFieldValue "FDecimal21", jxse '����˰��
    m_BillInterface.SetFieldValue "FDecimal23", qty '�������
    m_BillInterface.DoAction "FDecimal14"
    m_BillInterface.DoAction "FCheckBox"
    '�ļ��۸�
    Dim z01, z02, z03, z04, z05 As Double
'    If xsdj1 > 0 Then
'        z01 = xsdj1 / (1 + taxrate / 100)
'        z02 = z01 * taxrate / 100 - jxse
'        z03 = z02 * 0.12 + z01 * (0.001 + 0.8 * 0.0003)
'        z04 = z01 - ((YCLCB + BZWCB) * 1.005 + xiaoji + BCPJGF + ccf) * 1.021 - z03
'        z05 = z04 / z01 * 100
'        m_BillInterface.SetFieldValue "FDecimal8", xsdj1 '���ۼ۸�˰
'        m_BillInterface.SetFieldValue "FDecimal9", Round(z01, 2) '���ۼ۸񲻺�˰
'        m_BillInterface.SetFieldValue "FDecimal10", z02 'Ӧ����ֵ˰
'        m_BillInterface.SetFieldValue "FDecimal11", Round(z03, 2) '����˰
'        m_BillInterface.SetFieldValue "FDecimal12", Round(z04, 2) 'ë����
'        m_BillInterface.SetFieldValue "FDecimal13", Round(z05, 2) 'ë����
'        m_BillInterface.SetFieldValue "FDecimal25", xsdj1 / 1000 * qty '���
'    End If
    'Ԥ��۸�
'    If xsdj2 > 0 Then
'        z01 = xsdj2 / (1 + taxrate / 100)
'        z02 = z01 * taxrate / 100 - jxse
'        z03 = z02 * 0.12 + z01 * (0.001 + 0.8 * 0.0003)
'        z04 = z01 - ((YCLCB + BZWCB) * 1.005 + xiaoji + BCPJGF + ccf) * 1.021 - z03
'        z05 = z04 / z01 * 100
'        m_BillInterface.SetFieldValue "FDecimal14", xsdj2 '���ۼ۸�˰
'        m_BillInterface.SetFieldValue "FDecimal15", Round(z01, 2) '���ۼ۸񲻺�˰
'        m_BillInterface.SetFieldValue "FDecimal16", z02 '���ۼ۸�˰
'        m_BillInterface.SetFieldValue "FDecimal17", Round(z03, 2) '����˰
'        m_BillInterface.SetFieldValue "FDecimal18", Round(z04, 2) 'ë����
'        m_BillInterface.SetFieldValue "FDecimal19", Round(z05, 2) 'ë����
'        m_BillInterface.SetFieldValue "FDecimal20", xj '���
'    End If
    MsgBox "������ϣ�", vbOKOnly, oK3.LoadKDString("�����ʾ")
    
    GoTo ExitSub
Catch:
    Set dct = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, oK3.LoadKDString("�����ʾ")
    End If
ExitSub:
    Set oK3 = Nothing
    Set rs = Nothing
    End If
    
End Sub