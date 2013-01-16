VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Industry_CGSQD"
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

Private Sub m_BillTransfer_BillInitialize()
 
    'TODO: ���ڴ˴����Ӵ�����Ӧ�¼� BillInitialize
 
 
'*************** ��ʼ���ò˵� ***************
 
    m_BillTransfer.AddUserMenuItem "EXCEL����", ""
 
'*************** �������ò˵� ***************
 

End Sub
Private Function GetColIndex(ByVal ColHead As String, Optional iField As Long = 0) As Long
    On Error Resume Next
    Dim i As Long
    For i = 1 To m_BillTransfer.BillForm.vsEntrys.MaxCols
        If (m_BillTransfer.EntryCtl(i).HeadCaption = ColHead And iField = 0) Or (LCase(m_BillTransfer.EntryCtl(i).FieldName) = LCase(ColHead) And iField = 1) Then
            GetColIndex = i
            Exit For
        End If
    Next
End Function

Private Sub m_BillTransfer_UserMenuClick(ByVal Index As Long, ByVal Caption As String)
 
    'TODO: ���ڴ˴����Ӵ�����Ӧ�¼� UserMenuClick
    On Error GoTo Catch
 
    Select Case Caption
    Case "EXCEL����"
        '�˴����Ӵ��� EXCEL���� �˵������ Click �¼�
        If Not m_BillTransfer.SaveVect.Item(1).Value("BillChecked") Then
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
                i = 1
                If rs.EOF = False Then
                    If MsgBox("��ȷ��excel��ȷ��ʽ�������������Ϊ�������ϴ��롱����������λ������������;�������������ڡ�", vbYesNo) = vbYes Then
                    Do Until rs.EOF
                        m_BillTransfer.SetGridText i, GetColIndex("���ϴ���"), rs("���ϴ���").Value
                        m_BillTransfer.SetGridText i, GetColIndex("������λ����"), rs("������λ����").Value
                        m_BillTransfer.SetGridText i, GetColIndex("��;"), rs("��;").Value
                        m_BillTransfer.SetGridText i, GetColIndex("��������"), rs("��������").Value
                        Dim sqlstr As String
                        Dim rs2 As New ADODB.Recordset
                        sqlstr = "select t1.FCoefficient,t1.FNumber from t_MeasureUnit t1 where exists (select 1 from t_icitem t2 where t2.fnumber='" & rs("���ϴ���").Value & "' and t1.fitemid=t2.forderunitid)"
                        rs2.Open sqlstr, m_BillTransfer.Cnnstring, 0, 1
                        Dim qty As Double
                        qty = rs("������λ����").Value / rs2("FCoefficient")
                        m_BillTransfer.SetGridText i, GetColIndex("����"), qty
                        m_BillTransfer.SetGridText i, GetColIndex("��λ"), rs2("FNumber")
                        m_BillTransfer.BillForm.InsertRow 1, False
                        rs2.Close
                        Set rs2 = Nothing
                        i = i + 1
                        rs.MoveNext
                    Loop
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