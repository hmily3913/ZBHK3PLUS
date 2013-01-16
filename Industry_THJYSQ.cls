VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Industry_THJYSQ"
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

Private Sub m_BillTransfer_EndBillFormActive()
 
    'TODO: 请在此处添加代码响应事件 EndBillFormActive
    Dim EntryCtl As Variant
    Dim i, j As Long
    Dim vGrid As Object
    On Error GoTo Catch
    EntryCtl = m_BillTransfer.EntryCtl
    For i = 1 To UBound(m_BillTransfer.EntryCtl)
        If UCase(EntryCtl(i).FieldName) = UCase("Fitemid") Then
            Exit For
        End If
    Next
    Set vGrid = m_BillTransfer.Grid
    For j = 1 To UBound(m_BillTransfer.HeadCtl)
        If UCase(m_BillTransfer.HeadCtl(j).FieldName) = UCase("FCheckerID") Then
            Exit For
        End If
    Next
    
    '获取原单内容
    For z0 = 1 To UBound(m_BillTransfer.EntryCtl)
        If UCase(EntryCtl(z0).FieldName) = "FSOURCETRANTYPE" Then
            Exit For
        End If
    Next
    For z = 1 To UBound(m_BillTransfer.EntryCtl)
        If UCase(EntryCtl(z).FieldName) = "FSOURCEINTERID" Then
            Exit For
        End If
    Next
    For z1 = 1 To UBound(m_BillTransfer.EntryCtl)
        If UCase(EntryCtl(z1).FieldName) = "FSOURCEENTRYID" Then
            Exit For
        End If
    Next
    For z2 = 1 To UBound(m_BillTransfer.EntryCtl)
        If UCase(EntryCtl(z2).FieldName) = UCase("fsourcebillno") Then
            Exit For
        End If
    Next

    
    
    If m_BillTransfer.HeadCtl(j).VALE = "" Then
        With vGrid
            .Col = i
            .Col2 = i
            .Row = 1
            .Row2 = .MaxRows
            .BlockMode = True
            .Lock = False
            .BlockMode = False
        End With
    End If
    Exit Sub
Catch:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, oK3.LoadKDString("金蝶提示")
    End If
End Sub

Private Sub m_BillTransfer_RetEvents(ByVal Para As KFO.IDictionary)
 
    'TODO: 请在此处添加代码响应事件 RetEvents
End Sub
Private Sub m_BillTransfer_BeforeChangeMaterial(ByVal dctPara As KFO.IDictionary, bClearRow As Boolean)
 
    'TODO: 请在此处添加代码响应事件 BeforeChangeMaterial
    bClearRow = False
End Sub

