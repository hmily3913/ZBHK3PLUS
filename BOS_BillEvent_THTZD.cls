VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BillEvent_THTZD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is BillEvent Interface Class, made by K3BOSPLUGINSWIZAED"
Private Const App_Name = "ZBHK3PLUS"
Private Const ExePath = "\\10.0.0.5\gx\"
Private Const MidExeName = "FastCopy"

'���� BillEvent �ӿ�. ������е�����, �Դ�������¼�
Private WithEvents m_BillInterface  As BillEvent
Attribute m_BillInterface.VB_VarHelpID = -1
 
    '�汾��鼰���¹���
Private Sub ExeRefresh()
    '�����ĸ���ʱ�ַ�������
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    Dim s4 As String
    On Error Resume Next
    '������Ӧ�ó���MyApp.EXE��ȫ·�������� s1
    '�������н����FastCopy.exe��ȫ·�������� s3
    s1 = "TNT"
    If Len(App.Path) > 3 Then
        s1 = App.Path + "" + Trim(App_Name) + ".dll"
        s3 = App.Path + "" + MidExeName + ".EXE"
    Else
        s1 = App.Path + Trim(App_Name) + ".dll"
        s3 = App.Path + MidExeName + ".EXE"
    End If
    s4 = "TNT"
    '������Ӧ�ó���MyApp.EXE�ļ����޸�ʱ����� s4
    s4 = FileDateTime(s1)
    s2 = "TNT"
    '��������Ӧ�ó���MyApp.EXE�ļ����޸�ʱ����� s2
    s2 = FileDateTime(ExePath + App_Name + ".dll")
    If s2 = "TNT" Then
        MsgBox "û���ҵ����µĿ�ִ���ļ�:" + ExePath + App_Name + ".dll" + vbCrLf + vbCrLf + "ԭ��1:�������EXE�ķ��������߹���վû�д�;" + vbCrLf + "ԭ��2:�������EXE��·���������EXE�ļ�������;" + vbCrLf + "�뽫�����֪ͨ����Ա." + vbCrLf + vbCrLf + vbCrLf + "��ȷ����ť��,���������б���EXE�ļ�.", vbCritical, "��ʾ"
    End If
    If s2 = "TNT" Or s4 = "TNT" Then Exit Sub
    '���������Ӧ�ó���MyApp.EXE�ļ����޸�ʱ��,���ڱ���MyApp.EXE�ļ����޸�ʱ��
    'Ȼ�������б���MyApp.EXE ,�н�����˳���,�������¹��̽���.
    If CDate(s2) > CDate(s4) Then
        '�������ϵ��н����FastCopy.exe���Ƶ�����,�����ɷ�ֹ�������н����ʱ�޷����и���
        FileCopy ExePath + MidExeName + ".EXE", s3
        '�������н����FastCopy.exe ,�����µ�MyApp.EXE ���Ƶ�����
        s1 = Shell(s3 + " " + ExePath + "," + App_Name + ".EXE", vbNormalFocus)
        '����Ӧ�ó���MyApp.EXE ��ֹ����,�����Ѿ����µ�MyApp.EXE�޷����Ǳ��ص�MyApp.EXE .
    End If

End Sub

Public Sub Show(ByVal oBillInterface As Object)
' Call ExeRefresh
    'BillEvent �ӿ�ʵ��
    'ע��: �˷����������, �����޸�
    Set m_BillInterface = oBillInterface
 
End Sub

Private Sub Class_Terminate()
 
    '�ͷŽӿڶ���
    'ע��: �˷����������, �����޸�
    Set m_BillInterface = Nothing

End Sub
Private Sub m_BillInterface_BeforeSelBillShowFilter(ByRef dct As KFO.Dictionary)
    Dim sssss
    sssss = 1
End Sub
Private Sub m_BillInterface_AfterSelBillBeforeFillData(ByVal dcData As KFO.IDictionary, ByVal dctLink As KFO.IDictionary)
 
    'TODO: ���ڴ˴����Ӵ�����Ӧ�¼�
    On Error GoTo Catch
    Dim FIDStr As String
    Dim i As Integer
    For i = 1 To dcData("Page2").UBound - 1
'    dcData("Page2")(i)(m_BillInterface.TableInfo("map")("FClassID_SRC"))("FFLD")
        FIDStr = FIDStr & dcData("Page2")(i)(m_BillInterface.TableInfo("map")("FID_SRC"))("FFLD") & ","
    Next
    FIDStr = FIDStr & dcData("Page2")(i)(m_BillInterface.TableInfo("map")("FID_SRC"))("FFLD")
    Dim oK3  As K3ClassEvents.K3Lib
    Dim rs As ADODB.Recordset
    Set oK3 = m_BillInterface.K3Lib
    Set rs = oK3.GetData("select distinct FSupplyID from icstockbill where FTranType=21 and FInterID in (" & FIDStr & ")")
    If rs.RecordCount <> 1 Then
'        Set dcData = Nothing
'        m_BillInterface.SetBarStatus "mnuFileSave", "Enabled", False
'        m_BillInterface.MenuBar.Refresh
        MsgBox oK3.LoadKDString("ѡ��ĵ��ݱ���Ϊ��ͬ������λ��"), vbOKOnly + vbCritical, oK3.LoadKDString("�����ʾ")
'        dcData("Page2").RemoveAll
'    Else
'        m_BillInterface.SetBarStatus "mnuFileSave", "Enabled", True
'        m_BillInterface.MenuBar.Refresh
    End If
    GoTo ExitSub
Catch:
    Set dct = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, oK3.LoadKDString("�����ʾ")
    End If
ExitSub:
    Set oK3 = Nothing
    Set rs = Nothing
End Sub

Private Sub m_BillInterface_BeforeSave(ByRef bCancel As Boolean)

    On Error GoTo Catch
    Dim i As Integer
    Dim sqlstr As String
    Dim FIDStr As String
    
    With m_BillInterface
    For i = 1 To .Data("Page2").UBound
        If .GetFieldValue("FID_SRC", i, Enu_ValueType_FFLD) <> "" Then
            FIDStr = FIDStr & .GetFieldValue("FID_SRC", i, Enu_ValueType_FFLD) & ","
        End If
        
    Next
    End With
    If Len(FIDStr) > 0 Then
        FIDStr = Left(FIDStr, Len(FIDStr) - 1)
        sqlstr = "select FSupplyID from icstockbill where FTranType=21 and FInterID in (" & FIDStr & ")"
        Dim oK3  As K3ClassEvents.K3Lib
        Dim rs As ADODB.Recordset
        Set oK3 = m_BillInterface.K3Lib
        Set rs = oK3.GetData(sqlstr)
        If rs.RecordCount <> 1 Then
            bCancel = True
            MsgBox oK3.LoadKDString("ѡ��ĵ��ݱ���Ϊ��ͬ������λ������ʧ�ܣ�"), vbOKOnly + vbCritical, oK3.LoadKDString("�����ʾ")
        Else
            bCancel = False
        End If
    End If
    GoTo ExitSub
Catch:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, oK3.LoadKDString("�����ʾ")
    End If
ExitSub:
    Set oK3 = Nothing
    Set rs = Nothing
End Sub
Private Sub m_BillInterface_AfterSelBill(ByVal lSelBillType As Long)
    With m_BillInterface
    For i = 1 To .Data("Page2").UBound
        If .GetFieldValue("FID_SRC", i, Enu_ValueType_FFLD) <> "" Then
            FIDStr = FIDStr & .GetFieldValue("FID_SRC", i, Enu_ValueType_FFLD) & ","
        End If
        
    Next
    End With
End Sub

