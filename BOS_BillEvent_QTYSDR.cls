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

Private Sub m_BillInterface_MenuBarClick(ByVal BOSTool As K3ClassEvents.BOSTool, Cancel As Boolean)
 
    'TODO: 请在此处添加代码响应事件
    
    On Error GoTo Catch
    Select Case BOSTool.ToolName
    Case "mnuEXCELPIN"
        '此处添加处理mnuEXCELPIN 菜单对象的 Click 事件
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
                    If MsgBox("请确认excel正确格式，标题存在名称为：“客户代码”、“往来科目”、“单据日期”、“财务日期”、“摘要”、“金额”", vbYesNo) = vbYes Then
                    Do Until rs.EOF
                        m_BillInterface.InsertNewRowAndFill 2, i, "FBase", rs("客户代码"), "FBase1", rs("往来科目"), "FDate", rs("单据日期"), "FDate1", rs("财务日期"), "FText", IIf(IsNull(rs("摘要")), "", rs("摘要")), "FDecimal", rs("金额")
                        i = i + 1
                        rs.MoveNext
                    Loop
                    blnResult = m_BillInterface.BatchToDoActions("FBase", -1, -1)
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

Private Sub m_BillInterface_MenuBarInitialize(ByVal oMenuBar As K3ClassEvents.MenuBar)
 
    Dim oTool   As K3ClassEvents.BOSTool
    Dim oBand   As K3ClassEvents.BOSBand
 
'*************** 开始新增 BOS 菜单 ***************
 
    '新增 mnuEXCELPIN 菜单对象，并设置属性
    Set oTool = oMenuBar.BOSTools.Add("mnuEXCELPIN")
    With oTool
        .Caption = "EXCEL导入"
        .ToolTipText = "EXCEL导入"
        .Description = "EXCEL导入"
        .ShortcutKey = 0
        .ToolPicture = App.Path & "\excel.bmp" '直接加载文件
        .SetPicture 0, vbButtonFace '使用文件贴图
        .Visible = True
        .Enabled = True
        .BeginGroup = False
    End With
 
    Set oBand = oMenuBar.BOSBands("BandToolBar")
    oBand.BOSTools.InsertAfter "mnuDataRefesh", oTool   '将菜单对象插入指定工具栏
 
 
'*************** 结束新增 BOS 菜单 ***************
 

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
    
    If m_BillInterface.GetFieldValue("FDate2", , Enu_ValueType_FFLD) = "" Or m_BillInterface.GetFieldValue("FDate3", , Enu_ValueType_FFLD) = "" Then
        MsgBox "开始日期、截止日期不能为空，计算失败！", vbCritical, oK3.LoadKDString("金蝶提示")
        Exit Sub
    End If
    If m_BillInterface.GetFieldValue("FUser", , Enu_ValueType_FFLD) <> "" Then
        MsgBox "审核状态不允许计算，计算失败！", vbCritical, oK3.LoadKDString("金蝶提示")
        Exit Sub
    End If
    '删除现在分录
    Dim bln As Boolean
    bln = m_BillInterface.DeleteEntryData(2)
    
    strSQL = "select c.客户编号,sum(case when c.FQty>0 then c.运费金额 else c.FQty*isnull(e.FDecimal,0)/1000 end) as money,b.FName from vwicbill_8 c inner join t_Organization a on 客户编号=a.fnumber left join t_department b on a.FDepartment=b.FItemid left join z_YFJGEntry e on a.FDepartment=e.FBase and datediff(d,c.FDate,e.fdate)<=0 and datediff(d,c.FDate,e.FDate1)>=0 "
    strSQL = strSQL & " where FCheckFLag='※' and FCancellation='' and yfzffs=8610 and c.FDate>='" & m_BillInterface.GetFieldValue("FDate2", , Enu_ValueType_FFLD) & "' AND c.FDate<='" & m_BillInterface.GetFieldValue("FDate3", , Enu_ValueType_FFLD) & "' group by 客户编号,b.FName having sum(case when c.FQty>0 then c.运费金额 else c.FQty*isnull(e.FDecimal,0)/1000 end) <> 0 order by 客户编号"
    Set rs = oK3.GetData(strSQL)
    
    Dim i As Long
    i = 1
    While (Not rs.EOF)
        '批量插入行数据
        m_BillInterface.InsertNewRowAndFill 2, i, "FBase", rs("客户编号"), "FBase1", "2181.10", "FDate", Date, "FDate1", Date, "FText", "代扣" & rs("FName") & "运费", "FDecimal", rs("money")
        i = i + 1
        rs.MoveNext
    Wend
    
    MsgBox "读取完毕！", vbOKOnly, oK3.LoadKDString("金蝶提示")
    
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

