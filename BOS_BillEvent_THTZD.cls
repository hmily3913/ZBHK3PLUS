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

'定义 BillEvent 接口. 必须具有的声明, 以此来获得事件
Private WithEvents m_BillInterface  As BillEvent
Attribute m_BillInterface.VB_VarHelpID = -1
 
    '版本检查及更新过程
Private Sub ExeRefresh()
    '定义四个临时字符串变量
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    Dim s4 As String
    On Error Resume Next
    '将本地应用程序MyApp.EXE的全路径名存入 s1
    '将本地中介程序FastCopy.exe的全路径名存入 s3
    s1 = "TNT"
    If Len(App.Path) > 3 Then
        s1 = App.Path + "" + Trim(App_Name) + ".dll"
        s3 = App.Path + "" + MidExeName + ".EXE"
    Else
        s1 = App.Path + Trim(App_Name) + ".dll"
        s3 = App.Path + MidExeName + ".EXE"
    End If
    s4 = "TNT"
    '将本地应用程序MyApp.EXE文件的修改时间存入 s4
    s4 = FileDateTime(s1)
    s2 = "TNT"
    '将网络上应用程序MyApp.EXE文件的修改时间存入 s2
    s2 = FileDateTime(ExePath + App_Name + ".dll")
    If s2 = "TNT" Then
        MsgBox "没有找到最新的可执行文件:" + ExePath + App_Name + ".dll" + vbCrLf + vbCrLf + "原因1:存放最新EXE的服务器或者工作站没有打开;" + vbCrLf + "原因2:存放最新EXE的路径错误或者EXE文件不存在;" + vbCrLf + "请将此情况通知程序员." + vbCrLf + vbCrLf + vbCrLf + "按确定按钮后,将继续运行本地EXE文件.", vbCritical, "提示"
    End If
    If s2 = "TNT" Or s4 = "TNT" Then Exit Sub
    '如果网络上应用程序MyApp.EXE文件的修改时间,大于本地MyApp.EXE文件的修改时间
    '然后再运行本地MyApp.EXE ,中介程序退出后,整个更新过程结束.
    If CDate(s2) > CDate(s4) Then
        '将网络上的中介程序FastCopy.exe复制到本地,这样可防止本地无中介程序时无法进行更新
        FileCopy ExePath + MidExeName + ".EXE", s3
        '则运行中介程序FastCopy.exe ,将最新的MyApp.EXE 复制到本地
        s1 = Shell(s3 + " " + ExePath + "," + App_Name + ".EXE", vbNormalFocus)
        '本地应用程序MyApp.EXE 终止运行,否则已经更新的MyApp.EXE无法覆盖本地的MyApp.EXE .
    End If

End Sub

Public Sub Show(ByVal oBillInterface As Object)
' Call ExeRefresh
    'BillEvent 接口实现
    '注意: 此方法必须存在, 请勿修改
    Set m_BillInterface = oBillInterface
 
End Sub

Private Sub Class_Terminate()
 
    '释放接口对象
    '注意: 此方法必须存在, 请勿修改
    Set m_BillInterface = Nothing

End Sub
Private Sub m_BillInterface_BeforeSelBillShowFilter(ByRef dct As KFO.Dictionary)
    Dim sssss
    sssss = 1
End Sub
Private Sub m_BillInterface_AfterSelBillBeforeFillData(ByVal dcData As KFO.IDictionary, ByVal dctLink As KFO.IDictionary)
 
    'TODO: 请在此处添加代码响应事件
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
        MsgBox oK3.LoadKDString("选择的单据必须为相同购货单位！"), vbOKOnly + vbCritical, oK3.LoadKDString("金蝶提示")
'        dcData("Page2").RemoveAll
'    Else
'        m_BillInterface.SetBarStatus "mnuFileSave", "Enabled", True
'        m_BillInterface.MenuBar.Refresh
    End If
    GoTo ExitSub
Catch:
    Set dct = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, oK3.LoadKDString("金蝶提示")
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
            MsgBox oK3.LoadKDString("选择的单据必须为相同购货单位，保存失败！"), vbOKOnly + vbCritical, oK3.LoadKDString("金蝶提示")
        Else
            bCancel = False
        End If
    End If
    GoTo ExitSub
Catch:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, oK3.LoadKDString("金蝶提示")
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


