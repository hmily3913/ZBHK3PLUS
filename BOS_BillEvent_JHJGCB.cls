VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BillEvent_JHJGCB"
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
                    If MsgBox("请确认excel正确格式，标题存在名称为：“物料长代码”、“吨位直接人工”、“激光刻码”、“电费”、“设备折旧费”、“质检费”、“制造费用额”、“仓储费”", vbYesNo) = vbYes Then
'                    If MsgBox("请确认excel正确格式，标题存在名称为：“物料长代码”、“机器包装”、“手工包装”、“手工贴标费”、“激光刻码”、“生产工时”、“价值工时”、“工时价格”、“吨位工资价格”、“电费”、“设备折旧费”、“质检费”、“生产工时”、“单位工时分摊额”、“制造费用额”、“仓储费”", vbYesNo) = vbYes Then
                    Do Until rs.EOF
                        m_BillInterface.InsertNewRowAndFill 2, i, "FBase", rs("物料长代码"), "FDecimal", rs("吨位直接人工"), "FDecimal12", rs("激光刻码"), "FDecimal1", rs("电费"), "FDecimal13", rs("设备折旧费"), "FDecimal4", rs("质检费"), "FDecimal3", rs("制造费用额"), "FDecimal5", rs("仓储费")
'                        m_BillInterface.InsertNewRowAndFill 2, i, "FBase", rs("物料长代码"), "FDecimal8", rs("机器包装"), "FDecimal11", rs("手工包装"), "FDecimal2", rs("手工贴标费"), "FDecimal12", rs("激光刻码"), "FDecimal10", rs("生产工时"), "FDecimal7", rs("价值工时"), "FDecimal9", rs("工时价格"), "FDecimal", rs("吨位工资价格"), "FDecimal1", rs("电费"), "FDecimal13", rs("设备折旧费"), "FDecimal4", rs("质检费"), "FDecimal15", rs("生产工时"), "FDecimal14", rs("单位工时分摊额"), "FDecimal3", rs("制造费用额"), "FDecimal5", rs("仓储费")
                        i = i + 1
                        rs.MoveNext
                    Loop
                    blnResult = m_BillInterface.BatchToDoActions("FBase", -1, -1)
                    blnResult = m_BillInterface.BatchToDoActions("FDecimal6", -1, -1)
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
