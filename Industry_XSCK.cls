VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Industry_XSCK"
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

Private Sub m_BillTransfer_BeforeSave(ByVal bNew As Boolean, ReturnCode As Long)
 
    'TODO: 请在此处添加代码响应事件 BeforeSave
    
    Dim vsEntrys As fpSpread
    Dim EntryCtl As Variant
    Dim FSourceInterId As Variant
    Dim FSourceTranType As Variant
    Dim FSOURCEENTRYID As Variant
    Dim fauxqty, FUnitID As Variant
    Dim fscid As String
    Dim i As Long
    Dim z As Long
    Dim z1, z2, z3, z4, z5, ifunitid As Long
    Dim j As Long
    Dim sum As Double
    Dim vGrid As Object
    
    sum = 0
    EntryCtl = m_BillTransfer.EntryCtl
    For i = 1 To UBound(m_BillTransfer.EntryCtl)
        If UCase(EntryCtl(i).FieldName) = "FSOURCETRANTYPE" Then
            Exit For
        End If
    Next
    For z = 1 To UBound(m_BillTransfer.EntryCtl)
        If UCase(EntryCtl(z).FieldName) = "FSOURCEINTERID" Then
            Exit For
        End If
    Next
    For ifunitid = 1 To UBound(m_BillTransfer.EntryCtl)
        If LCase(EntryCtl(ifunitid).FieldName) = "funitid" Then
            Exit For
        End If
    Next
    For z1 = 1 To UBound(m_BillTransfer.EntryCtl)
        If UCase(EntryCtl(z1).FieldName) = "FSOURCEENTRYID" Then
            Exit For
        End If
    Next
    For z2 = 1 To UBound(m_BillTransfer.EntryCtl)
        If LCase(EntryCtl(z2).FieldName) = "fauxprice" Then
            Exit For
        End If
    Next
    For z3 = 1 To UBound(m_BillTransfer.EntryCtl)
        If LCase(EntryCtl(z3).FieldName) = "famount" Then
            Exit For
        End If
    Next
    For z4 = 1 To UBound(m_BillTransfer.EntryCtl)
        If LCase(EntryCtl(z4).FieldName) = "fconsignprice" Then
            Exit For
        End If
    Next
    For z5 = 1 To UBound(m_BillTransfer.EntryCtl)
        If LCase(EntryCtl(z5).FieldName) = "fconsignamount" Then
            Exit For
        End If
    Next
    For z6 = 1 To UBound(m_BillTransfer.EntryCtl)
        If LCase(EntryCtl(z6).FieldName) = "fauxqty" Then
            Exit For
        End If
    Next
    Set vGrid = m_BillTransfer.Grid
    
'    Dim objAPP As Object
    Dim rs As ADODB.Recordset
'    Set conn = New ADODB.Connection
    On Error GoTo Catch
'    Set objAPP = CreateObject("K3MAppconnection.AppConnection")
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
    
    With vGrid
    For j = 1 To .MaxRows
        .GetText i, j, FSourceTranType
        .GetText z, j, FSourceInterId
        .GetText z1, j, FSOURCEENTRYID
        .GetText z6, j, fauxqty
        .GetText ifunitid, j, FUnitID
        If FSourceTranType = "704" Then
'            conn.Open m_BillTransfer.Cnnstring
            sqlstr = "select b.FEntrySelfT0434,b.FEntrySelfT0435,b.FEntrySelfT0436,b.FEntrySelfT0437,f.FCoefficient from QMSaleCKRequest a,QMSaleCKRequestEntry b,ICSaleQcBill c,t_submessage d,t_icitembase e,t_measureUnit f where b.fitemid=e.fitemid and e.Funitgroupid=f.Funitgroupid and f.fname='" & FUnitID & "' and a.Finterid=b.Finterid and a.FTranType=704 and a.FInterID=" & FSourceInterId & " and b.FEntryID=" & FSOURCEENTRYID & " and c.FTranType=707 and c.FInStockInterID=b.FinterID and c.FSerialID=b.FEntryID and c.FHeadSelft0752=d.FInterid and d.FName='符合'"
            Set rs = New ADODB.Recordset
            rs.Open sqlstr, m_BillTransfer.Cnnstring, 0, 1
            If rs.EOF Then
                ReturnCode = -1
                MsgBox "分录：" & j & " 对应的原检验申请单，检验结果为不符合退货描述，不允许生成红字销售出库，保存失败！"
                Exit Sub
            Else
                Dim famount, fconsignamount, price, conprice As Double
                price = CDbl(rs("FEntrySelfT0437")) * CDbl(rs("FCoefficient"))
                conprice = CDbl(rs("FEntrySelfT0436")) * CDbl(rs("FCoefficient"))
                famount = price * CDbl(fauxqty)
                fconsignamount = conprice * CDbl(fauxqty)
                .SetText z2, j, price
                .SetText z3, j, famount
                .SetText z4, j, conprice
                .SetText z5, j, fconsignamount
            End If
        ElseIf FSourceTranType = "" Then
            Exit For
        End If
    Next
    End With
    
    GoTo ExitSub
Catch:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, oK3.LoadKDString("金蝶提示")
    End If
ExitSub:
    Set rs = Nothing
'    Set conn = Nothing
End Sub
