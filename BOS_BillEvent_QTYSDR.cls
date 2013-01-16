VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BillEvent_QTYSDR"
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

Private Sub m_BillInterface_MenuBarClick(ByVal BOSTool As K3ClassEvents.BOSTool, Cancel As Boolean)
 
    'TODO: ���ڴ˴����Ӵ�����Ӧ�¼�
    
    On Error GoTo Catch
    Select Case BOSTool.ToolName
    Case "mnuEXCELPIN"
        '�˴����Ӵ���mnuEXCELPIN �˵������ Click �¼�
        If m_BillInterface.BillStatus = Enu_BillStatusExt_New Then
            Dim strDir     As String
            SelectFileForm.Show 1
            strDir = SelectFileForm.Text1.Text
            Unload SelectFileForm
            If Dir(strDir, vbDirectory) <> "" Then
                Dim adoConnection  As New ADODB.Connection
                Dim sSQL As String
                Dim rs As New ADODB.Recordset
                
                adoConnection.Open "Data Provider=MSDASQL.1;driver=Microsoft Excel Driver (*.xls);DBQ=" & strDir
                sSQL = "select * from [sheet1$]"
                rs.Open sSQL, adoConnection, adOpenDynamic, adLockPessimistic
                Dim i As Long
                Dim blnResult As Boolean
                i = 1
                If rs.EOF = False Then
                    If MsgBox("��ȷ��excel��ȷ��ʽ�������������Ϊ�����ͻ����롱����������Ŀ�������������ڡ������������ڡ�����ժҪ��������", vbYesNo) = vbYes Then
                    Do Until rs.EOF
                        m_BillInterface.InsertNewRowAndFill 2, i, "FBase", rs("�ͻ�����"), "FBase1", rs("������Ŀ"), "FDate", rs("��������"), "FDate1", rs("��������"), "FText", IIf(IsNull(rs("ժҪ")), "", rs("ժҪ")), "FDecimal", rs("���")
                        i = i + 1
                        rs.MoveNext
                    Loop
                    blnResult = m_BillInterface.BatchToDoActions("FBase", -1, -1)
                    MsgBox "������ϣ�", vbOKOnly, "�����ʾ"
                    End If
                    
                Else
                    MsgBox ("excel�Ĺ��������Ʊ���Ϊsheet1")
                    Exit Sub
                End If
                rs.Close
                Set rs = Nothing
                adoConnection.Close
                Set adoConnection = Nothing

            Else
                MsgBox ("�ļ������ڣ�������ѡ��")
                Exit Sub
            End If
        Else
            MsgBox "����״̬Ϊ����ʱ���������룬��ǰ״̬���������룡", vbExclamation, "EXCEL����"
            Exit Sub
        End If
    Case Else
    End Select
    Exit Sub
Catch:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "�����ʾ"
    End If

End Sub

Private Sub m_BillInterface_MenuBarInitialize(ByVal oMenuBar As K3ClassEvents.MenuBar)
 
    Dim oTool   As K3ClassEvents.BOSTool
    Dim oBand   As K3ClassEvents.BOSBand
 
'*************** ��ʼ���� BOS �˵� ***************
 
    '���� mnuEXCELPIN �˵����󣬲���������
    Set oTool = oMenuBar.BOSTools.Add("mnuEXCELPIN")
    With oTool
        .Caption = "EXCEL����"
        .ToolTipText = "EXCEL����"
        .Description = "EXCEL����"
        .ShortcutKey = 0
        .ToolPicture = App.Path & "\excel.bmp" 'ֱ�Ӽ����ļ�
        .SetPicture 0, vbButtonFace 'ʹ���ļ���ͼ
        .Visible = True
        .Enabled = True
        .BeginGroup = False
    End With
 
    Set oBand = oMenuBar.BOSBands("BandToolBar")
    oBand.BOSTools.InsertAfter "mnuDataRefesh", oTool   '���˵��������ָ��������
 
 
'*************** �������� BOS �˵� ***************
 

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
    
    If m_BillInterface.GetFieldValue("FDate2", , Enu_ValueType_FFLD) = "" Or m_BillInterface.GetFieldValue("FDate3", , Enu_ValueType_FFLD) = "" Then
        MsgBox "��ʼ���ڡ���ֹ���ڲ���Ϊ�գ�����ʧ�ܣ�", vbCritical, oK3.LoadKDString("�����ʾ")
        Exit Sub
    End If
    If m_BillInterface.GetFieldValue("FUser", , Enu_ValueType_FFLD) <> "" Then
        MsgBox "���״̬���������㣬����ʧ�ܣ�", vbCritical, oK3.LoadKDString("�����ʾ")
        Exit Sub
    End If
    'ɾ�����ڷ�¼
    Dim bln As Boolean
    bln = m_BillInterface.DeleteEntryData(2)
    
    strSQL = "select c.�ͻ����,sum(case when c.FQty>0 then c.�˷ѽ�� else c.FQty*isnull(e.FDecimal,0)/1000 end) as money,b.FName from vwicbill_8 c inner join t_Organization a on �ͻ����=a.fnumber left join t_department b on a.FDepartment=b.FItemid left join z_YFJGEntry e on a.FDepartment=e.FBase and datediff(d,c.FDate,e.fdate)<=0 and datediff(d,c.FDate,e.FDate1)>=0 "
    strSQL = strSQL & " where FCheckFLag='��' and FCancellation='' and yfzffs=8610 and c.FDate>='" & m_BillInterface.GetFieldValue("FDate2", , Enu_ValueType_FFLD) & "' AND c.FDate<='" & m_BillInterface.GetFieldValue("FDate3", , Enu_ValueType_FFLD) & "' group by �ͻ����,b.FName having sum(case when c.FQty>0 then c.�˷ѽ�� else c.FQty*isnull(e.FDecimal,0)/1000 end) <> 0 order by �ͻ����"
    Set rs = oK3.GetData(strSQL)
    
    Dim i As Long
    i = 1
    While (Not rs.EOF)
        '��������������
        m_BillInterface.InsertNewRowAndFill 2, i, "FBase", rs("�ͻ����"), "FBase1", "2181.10", "FDate", Date, "FDate1", Date, "FText", "����" & rs("FName") & "�˷�", "FDecimal", rs("money")
        i = i + 1
        rs.MoveNext
    Wend
    
    MsgBox "��ȡ��ϣ�", vbOKOnly, oK3.LoadKDString("�����ʾ")
    
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
