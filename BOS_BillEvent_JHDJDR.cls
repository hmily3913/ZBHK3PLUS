VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BillEvent_JHDJDR"
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
                    If MsgBox("��ȷ��excel��ȷ��ʽ�������������Ϊ�������ϳ����롱�������ϵ��ۡ�����Ԥ��ϵ����", vbYesNo) = vbYes Then
                    Do Until rs.EOF
                        m_BillInterface.InsertNewRowAndFill 2, i, "FBase", rs("���ϳ�����"), "FDecimal", rs("���ϵ���"), "FDecimal2", rs("Ԥ��ϵ��")
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