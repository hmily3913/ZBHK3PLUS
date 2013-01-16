VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MiddleAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This is IndustryMiddleWareEvents Interface Class, made by K3BOSPLUGINSWIZAED"
 
Public Function HookInvoke(ByVal EventID As Long, dctParams As KFO.Dictionary) As Long
    Dim objAPP As Object
    Dim adoRS As ADODB.Recordset
    Dim strSQL As String
    
    Set objAPP = CreateObject("K3MAppconnection.AppConnection")
    
    strSQL = "insert into xxpt.dbo.te (a) values (1)"
    Set adoRS = objAPP.getanyrecordset(MMTS.PropsString, strSQL)
    Set adoRS = Nothing
    Set objAPP = Nothing
    
dctParams("ErrCode") = 33 '可以自定义
dctParams("ErrSource") = "zzffReDevPro.clsMid审核判断组件"
dctParams("ErrDescription") = "zzff二次开发提示：单价为0，不能审核"
HookInvoke = 3
Exit Function
    '定义插件对象接口. 必须具有的声明, 以此来获得事件
    '注意: 此方法必须存在, 请勿修改
    Select Case EventID
        Case 300010 '单据关联反写后事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'DestBillInterID        | 目标单据的内码
            'DestTranType           | 目标单据的事务类型
            'StoreSwitch            | 反写方式。值为 1 时为正向反写；否则为 反向反写
            'WithNoCheck            | 是否在反写前检查
            'WriteBackTemplate      | 反写模版对象
            '--------------------------------------------------------

        Case 300009 '单据关联反写前事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'DestBillInterID        | 目标单据的内码
            'DestTranType           | 目标单据的事务类型
            'StoreSwitch            | 反写方式。值为 1 时为正向反写；否则为 反向反写
            'WithNoCheck            | 是否在反写前检查
            '--------------------------------------------------------

        Case 300008 '关闭/反关闭单据后事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'BillInterID            | 单据的内码
            'TableHeadName          | 表头表名称
            'CloseSwitch            | 值为 0
            '--------------------------------------------------------

        Case 300007 '关闭/反关闭单据前事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'BillInterID            | 单据的内码
            'TableHeadName          | 表头表名称
            'CloseSwitch            | 值为 0
            '--------------------------------------------------------

        Case 300006 '作废/反作废单据后事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'BillInterID            | 单据的内码
            'TranType               | 单据的事务类型
            'OperateCode            | 操作码:   16------反作废
            '                       |           32------作废
            'TableHeadName          | 表头表名称
            'CancelSwitch           | 值为 0  代表为 作废 操作；否则为 反作废
            'retCode                | 返回代码
            'retDesc                | 返回描述
            'ReCancel               | 作废/反作废控制代码
            'ReCancel2              | 作废/反作废控制代码
            'RetDesc2               | 返回描述
            'CheckItemInfo          | 返回的信息（信用检查使用）
            '--------------------------------------------------------

        Case 300005 '作废/反作废单据前事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'BillInterID            | 单据的内码
            'TranType               | 单据的事务类型
            'OperateCode            | 操作码:   16------反作废
            '                       |           32------作废
            'TableHeadName          | 表头表名称
            'CancelSwitch           | 值为 0  代表为 作废 操作；否则为 反作废
            'retCode                | 返回代码
            'retDesc                | 返回描述
            'ReCancel               | 作废/反作废控制代码
            'ReCancel2              | 作废/反作废控制代码
            'RetDesc2               | 返回描述
            'CheckItemInfo          | 返回的信息（信用检查使用）
            '--------------------------------------------------------

        Case 300004 '删除单据后事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'InterID                | 单据的内码
            'TranType               | 单据的事务类型
            'OperateCode            | 操作码。固定值。值为 8
            'ListID                 | 序时簿ID
            'BaseInfo               | 连接信息
            'BillClsID              | 单据类别ID
            'TemplateID             | 序时簿模版ID
            'ErrInfo                | 返回的错误信息
            'ErrInfo2               | 返回的错误信息
            'Return                 | 返回的错误代码
            'ReCheck                | 检查控制代码
            'ReCheck2               | 检查控制代码
            'CheckItemInfo          | 返回的信息（信用检查使用）
            '--------------------------------------------------------

        Case 300003 '删除单据前事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'InterID                | 单据的内码
            'TranType               | 单据的事务类型
            'OperateCode            | 操作码。固定值。值为 8
            'ListID                 | 序时簿ID
            'BaseInfo               | 连接信息
            'BillClsID              | 单据类别ID
            'TemplateID             | 序时簿模版ID
            'ErrInfo                | 返回的错误信息
            'ErrInfo2               | 返回的错误信息
            'Return                 | 返回的错误代码
            'ReCheck                | 检查控制代码
            'ReCheck2               | 检查控制代码
            'CheckItemInfo          | 返回的信息（信用控制使用）
            '--------------------------------------------------------

        Case 300002 '序时簿复制后事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'BillInterID            | 单据的内码
            'TranType               | 单据的事务类型
            'OperateCode            | 操作码。固定值。值为 64
            'ErrCode                | 错误代码。固定值。值为 -300002
            'ErrSource              | 错误来源。固定值。值为 "K3ListServer.clsLitUpdate.CopyBill.HookInvoke"
            'ErrDescription         | 错误描述。固定值。值为 "HookInvoke Error"
            'BillStatus             | 系统设置：关账状态。 0 - 未关账；1 - 已关账
            'ICClosed               | 系统设置：业务系统是否完成初始化
            'UpStockWhenSave        | 系统设置：库存更新控制。True - 保存更新；False - 审核更新
            'FFinishInitAP          | 应收应付是否启用。0 - 未启用
            'BillDict               | 单据对象
            '--------------------------------------------------------

        Case 300001 '序时簿复制前事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'BillInterID            | 单据的内码
            'TranType               | 单据的事务类型
            'OperateCode            | 操作码。固定值。值为 64
            'ErrCode                | 错误代码。固定值。值为 -300001
            'ErrSource              | 错误来源。固定值。值为 "K3ListServer.clsLitUpdate.CopyBill.HookInvoke"
            'ErrDescription         | 错误描述。固定值。值为 "HookInvoke Error"
            'BillStatus             | 系统设置：关账状态。 0 - 未关账；1 - 已关账
            'ICClosed               | 系统设置：业务系统是否完成初始化
            'UpStockWhenSave        | 系统设置：库存更新控制。True - 保存更新；False - 审核更新
            'FFinishInitAP          | 应收应付是否启用。0 - 未启用
            'BillDict               | 单据对象
            '--------------------------------------------------------

        Case 200003 '审核退出前事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'InterID                | 单据的内码
            'TranType               | 单据的事务类型
            'OperateDelete          | 默认值为0
            'ROB                    | 是否是红单
            'OperateCode            | 操作码:   1------审核
            '                       |           2------反审核
            '                       |           4------保存
            '                       |           8------删除
            '                       |           16-----反作废
            '                       |           32-----作废
            '--------------------------------------------------------

        Case 200002 '审核反写事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'InterID                | 单据的内码
            'TranType               | 单据的事务类型
            'OperateDelete          | 默认值为0
            'ROB                    | 是否是红单
            'OperateCode            | 操作码:   1------审核
            '                       |           2------反审核
            '                       |           4------保存
            '                       |           8------删除
            '                       |           16-----反作废
            '                       |           32-----作废
            '--------------------------------------------------------

        Case 200001 '审核前事件
            '--------------------------------------------------------
 
            'dctParams 值域         | 含义
            '--------------------------------------------------------
            'sDsn                   | 连接字符串
            'InterID                | 单据的内码
            'TranType               | 单据的事务类型
            'OperateDelete          | 默认值为0
            'ROB                    | 是否是红单
            'OperateCode            | 操作码:   1------审核
            '                       |           2------反审核
            '                       |           4------保存
            '                       |           8------删除
            '                       |           16-----反作废
            '                       |           32-----作废
            '--------------------------------------------------------
            
            '请在这里添加处理 '审核前' 代码

 
    Case Else
    End Select
 
End Function
Public Function EndSave(ByVal sDsn As String, vctSave As KFO.Vector, sReturnMsg As String) As Boolean
    '-----------------------------------------------------------------------
    '参数                   | 含义
    '-----------------------------------------------------------------------
    'sDsn 参数              | 连接字符串
    'sReturnMsg 参数        | 返回的错误信息
    'vctSave                | 保存时信息包。值域构成如下
    '-----------------------------------------------------------------------
    
    '-----------------------------------------------------------------------
    'vctSave 值域           | 含义
    '-----------------------------------------------------------------------
    'FInterID               | 单据的内码
    'FTransType             | 单据的事务类型
    'IsRedBill              | 当前单据是否是红单
    'BillChecked            | 当前单据是否审核
    'CurrentUserRight       | 当前用户是否有权限
    'ActiveCtlValueChanged  | 当前单据是否改变
    'BeforeCloseStockBill   | 当前单据是否启用期前仓存单据(暂估入库，销售出库)
    'ICClosed               | 当前是否关账
    '-----------------------------------------------------------------------
    
    On Error GoTo Catch
    '请在这里添加处理 EndSave 代码
    
    
    '如果正常完成，需要返回 True
    EndSave = True
    Exit Function
    
Catch:
    '如果出现错误，需要返回 False，以及错误描述
    EndSave = False
    sReturnMsg = Err.Description
End Function
 
Public Function BeginSave(ByVal sDsn As String, vctSave As KFO.Vector, sReturnMsg As String) As Boolean
    '-----------------------------------------------------------------------
    '参数                   | 含义
    '-----------------------------------------------------------------------
    'sDsn 参数              | 连接字符串
    'sReturnMsg 参数        | 返回的错误信息
    'vctSave                | 保存时信息包。值域构成如下
    '-----------------------------------------------------------------------
    
    '-----------------------------------------------------------------------
    'vctSave 值域           | 含义
    '-----------------------------------------------------------------------
    'FInterID               | 单据的内码
    'FTransType             | 单据的事务类型
    'IsRedBill              | 当前单据是否是红单
    'BillChecked            | 当前单据是否审核
    'CurrentUserRight       | 当前用户是否有权限
    'ActiveCtlValueChanged  | 当前单据是否改变
    'BeforeCloseStockBill   | 当前单据是否启用期前仓存单据(暂估入库，销售出库)
    'ICClosed               | 当前是否关账
    '-----------------------------------------------------------------------
    
    On Error GoTo Catch
    '请在这里添加处理 BeginSave 代码
    
    
    '如果正常完成，需要返回 True
    BeginSave = True
    Exit Function
    
Catch:
    '如果出现错误，需要返回 False，以及错误描述
    BeginSave = False
    sReturnMsg = Err.Description
End Function
 
