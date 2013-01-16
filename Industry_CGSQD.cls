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

Private Sub m_BillTransfer_BillInitialize()
 
    'TODO: 请在此处添加代码响应事件 BillInitialize
 
 
'*************** 开始设置菜单 ***************
 
    m_BillTransfer.AddUserMenuItem "EXCEL导入", ""
 
'*************** 结束设置菜单 ***************
 

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
 
    'TODO: 请在此处添加代码响应事件 UserMenuClick
    On Error GoTo Catch
 
    Select Case Caption
    Case "EXCEL导入"
        '此处添加处理 EXCEL导入 菜单对象的 Click 事件
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
                    If MsgBox("请确认excel正确格式，标题存在名称为：“物料代码”、“基本单位数量”、“用途”、“到货日期”", vbYesNo) = vbYes Then
                    Do Until rs.EOF
                        m_BillTransfer.SetGridText i, GetColIndex("物料代码"), rs("物料代码").Value
                        m_BillTransfer.SetGridText i, GetColIndex("基本单位数量"), rs("基本单位数量").Value
                        m_BillTransfer.SetGridText i, GetColIndex("用途"), rs("用途").Value
                        m_BillTransfer.SetGridText i, GetColIndex("到货日期"), rs("到货日期").Value
                        Dim sqlstr As String
                        Dim rs2 As New ADODB.Recordset
                        sqlstr = "select t1.FCoefficient,t1.FNumber from t_MeasureUnit t1 where exists (select 1 from t_icitem t2 where t2.fnumber='" & rs("物料代码").Value & "' and t1.fitemid=t2.forderunitid)"
                        rs2.Open sqlstr, m_BillTransfer.Cnnstring, 0, 1
                        Dim qty As Double
                        qty = rs("基本单位数量").Value / rs2("FCoefficient")
                        m_BillTransfer.SetGridText i, GetColIndex("数量"), qty
                        m_BillTransfer.SetGridText i, GetColIndex("单位"), rs2("FNumber")
                        m_BillTransfer.BillForm.InsertRow 1, False
                        rs2.Close
                        Set rs2 = Nothing
                        i = i + 1
                        rs.MoveNext
                    Loop
                    MsgBox "导入完毕！", vbOKOnly, "金蝶提示"
                    End If
                    
                Else
                    MsgBox ("excel的工作表名称必须为sheet1")
                    Exit Sub
                End If
                rs.Close
                Set rs = Nothing
                adoConnection.Close
                Set adoConnection = Nothing

            Else
                MsgBox ("文件不存在，请重新选择！")
                Exit Sub
            End If
        Else
            MsgBox "单据状态为新增时才允许导入，当前状态不允许导入！", vbExclamation, "EXCEL导入"
            Exit Sub
        End If

    Case Else
    End Select
    Exit Sub
Catch:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "金蝶提示"
    End If

End Sub
