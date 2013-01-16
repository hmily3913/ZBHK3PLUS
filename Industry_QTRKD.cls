VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Industry_QTRKD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
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

Private Function GetHeadIndex(ByVal ColHead As String, Optional iField As Long = 0) As Long
    On Error Resume Next
    Dim i As Long

    For i = 1 To UBound(m_BillTransfer.HeadCtl)
        If (m_BillTransfer.HeadCtl(i).Caption = ColHead And iField = 0) Or (LCase(m_BillTransfer.HeadCtl(i).FieldName) = LCase(ColHead) And iField = 1) Then
            GetHeadIndex = i
            Exit For
        End If
    Next

End Function
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

Private Sub m_BillTransfer_BeforeSave(ByVal bNew As Boolean, ReturnCode As Long)
 
     'TODO: ���ڴ˴����Ӵ�����Ӧ�¼� BeforeSave
    Dim cklx, xm As String
    cklx = m_BillTransfer.GetHeadNumber(GetHeadIndex("�������:")) '�з���Ʒ
    xm = m_BillTransfer.GetHeadText(GetHeadIndex("��Ŀ"))
    If cklx = "010" And xm = "" Then
        ReturnCode = -1
        MsgBox "�з�����ʱ����Ŀ���Ʋ���Ϊ�գ�"
    End If
End Sub
