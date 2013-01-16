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
    
dctParams("ErrCode") = 33 '�����Զ���
dctParams("ErrSource") = "zzffReDevPro.clsMid����ж����"
dctParams("ErrDescription") = "zzff���ο�����ʾ������Ϊ0���������"
HookInvoke = 3
Exit Function
    '����������ӿ�. ������е�����, �Դ�������¼�
    'ע��: �˷����������, �����޸�
    Select Case EventID
        Case 300010 '���ݹ�����д���¼�
            '--------------------------------------------------------
 
            'dctParams ֵ��         | ����
            '--------------------------------------------------------
            'sDsn                   | �����ַ���
            'DestBillInterID        | Ŀ�굥�ݵ�����
            'DestTranType           | Ŀ�굥�ݵ���������
            'StoreSwitch            | ��д��ʽ��ֵΪ 1 ʱΪ����д������Ϊ ����д
            'WithNoCheck            | �Ƿ��ڷ�дǰ���
            'WriteBackTemplate      | ��дģ�����
            '--------------------------------------------------------

        Case 300009 '���ݹ�����дǰ�¼�
            '--------------------------------------------------------
 
            'dctParams ֵ��         | ����
            '--------------------------------------------------------
            'sDsn                   | �����ַ���
            'DestBillInterID        | Ŀ�굥�ݵ�����
            'DestTranType           | Ŀ�굥�ݵ���������
            'StoreSwitch            | ��д��ʽ��ֵΪ 1 ʱΪ����д������Ϊ ����д
            'WithNoCheck            | �Ƿ��ڷ�дǰ���
            '--------------------------------------------------------

        Case 300008 '�ر�/���رյ��ݺ��¼�
            '--------------------------------------------------------
 
            'dctParams ֵ��         | ����
            '--------------------------------------------------------
            'sDsn                   | �����ַ���
            'BillInterID            | ���ݵ�����
            'TableHeadName          | ��ͷ������
            'CloseSwitch            | ֵΪ 0
            '--------------------------------------------------------

        Case 300007 '�ر�/���رյ���ǰ�¼�
            '--------------------------------------------------------
 
            'dctParams ֵ��         | ����
            '--------------------------------------------------------
            'sDsn                   | �����ַ���
            'BillInterID            | ���ݵ�����
            'TableHeadName          | ��ͷ������
            'CloseSwitch            | ֵΪ 0
            '--------------------------------------------------------

        Case 300006 '����/�����ϵ��ݺ��¼�
            '--------------------------------------------------------
 
            'dctParams ֵ��         | ����
            '--------------------------------------------------------
            'sDsn                   | �����ַ���
            'BillInterID            | ���ݵ�����
            'TranType               | ���ݵ���������
            'OperateCode            | ������:   16------������
            '                       |           32------����
            'TableHeadName          | ��ͷ������
            'CancelSwitch           | ֵΪ 0  ����Ϊ ���� ����������Ϊ ������
            'retCode                | ���ش���
            'retDesc                | ��������
            'ReCancel               | ����/�����Ͽ��ƴ���
            'ReCancel2              | ����/�����Ͽ��ƴ���
            'RetDesc2               | ��������
            'CheckItemInfo          | ���ص���Ϣ�����ü��ʹ�ã�
            '--------------------------------------------------------

        Case 300005 '����/�����ϵ���ǰ�¼�
            '--------------------------------------------------------
 
            'dctParams ֵ��         | ����
            '--------------------------------------------------------
            'sDsn                   | �����ַ���
            'BillInterID            | ���ݵ�����
            'TranType               | ���ݵ���������
            'OperateCode            | ������:   16------������
            '                       |           32------����
            'TableHeadName          | ��ͷ������
            'CancelSwitch           | ֵΪ 0  ����Ϊ ���� ����������Ϊ ������
            'retCode                | ���ش���
            'retDesc                | ��������
            'ReCancel               | ����/�����Ͽ��ƴ���
            'ReCancel2              | ����/�����Ͽ��ƴ���
            'RetDesc2               | ��������
            'CheckItemInfo          | ���ص���Ϣ�����ü��ʹ�ã�
            '--------------------------------------------------------

        Case 300004 'ɾ�����ݺ��¼�
            '--------------------------------------------------------
 
            'dctParams ֵ��         | ����
            '--------------------------------------------------------
            'sDsn                   | �����ַ���
            'InterID                | ���ݵ�����
            'TranType               | ���ݵ���������
            'OperateCode            | �����롣�̶�ֵ��ֵΪ 8
            'ListID                 | ��ʱ��ID
            'BaseInfo               | ������Ϣ
            'BillClsID              | �������ID
            'TemplateID             | ��ʱ��ģ��ID
            'ErrInfo                | ���صĴ�����Ϣ
            'ErrInfo2               | ���صĴ�����Ϣ
            'Return                 | ���صĴ������
            'ReCheck                | �����ƴ���
            'ReCheck2               | �����ƴ���
            'CheckItemInfo          | ���ص���Ϣ�����ü��ʹ�ã�
            '--------------------------------------------------------

        Case 300003 'ɾ������ǰ�¼�
            '--------------------------------------------------------
 
            'dctParams ֵ��         | ����
            '--------------------------------------------------------
            'sDsn                   | �����ַ���
            'InterID                | ���ݵ�����
            'TranType               | ���ݵ���������
            'OperateCode            | �����롣�̶�ֵ��ֵΪ 8
            'ListID                 | ��ʱ��ID
            'BaseInfo               | ������Ϣ
            'BillClsID              | �������ID
            'TemplateID             | ��ʱ��ģ��ID
            'ErrInfo                | ���صĴ�����Ϣ
            'ErrInfo2               | ���صĴ�����Ϣ
            'Return                 | ���صĴ������
            'ReCheck                | �����ƴ���
            'ReCheck2               | �����ƴ���
            'CheckItemInfo          | ���ص���Ϣ�����ÿ���ʹ�ã�
            '--------------------------------------------------------

        Case 300002 '��ʱ�����ƺ��¼�
            '--------------------------------------------------------
 
            'dctParams ֵ��         | ����
            '--------------------------------------------------------
            'sDsn                   | �����ַ���
            'BillInterID            | ���ݵ�����
            'TranType               | ���ݵ���������
            'OperateCode            | �����롣�̶�ֵ��ֵΪ 64
            'ErrCode                | ������롣�̶�ֵ��ֵΪ -300002
            'ErrSource              | ������Դ���̶�ֵ��ֵΪ "K3ListServer.clsLitUpdate.CopyBill.HookInvoke"
            'ErrDescription         | �����������̶�ֵ��ֵΪ "HookInvoke Error"
            'BillStatus             | ϵͳ���ã�����״̬�� 0 - δ���ˣ�1 - �ѹ���
            'ICClosed               | ϵͳ���ã�ҵ��ϵͳ�Ƿ���ɳ�ʼ��
            'UpStockWhenSave        | ϵͳ���ã������¿��ơ�True - ������£�False - ��˸���
            'FFinishInitAP          | Ӧ��Ӧ���Ƿ����á�0 - δ����
            'BillDict               | ���ݶ���
            '--------------------------------------------------------

        Case 300001 '��ʱ������ǰ�¼�
            '--------------------------------------------------------
 
            'dctParams ֵ��         | ����
            '--------------------------------------------------------
            'sDsn                   | �����ַ���
            'BillInterID            | ���ݵ�����
            'TranType               | ���ݵ���������
            'OperateCode            | �����롣�̶�ֵ��ֵΪ 64
            'ErrCode                | ������롣�̶�ֵ��ֵΪ -300001
            'ErrSource              | ������Դ���̶�ֵ��ֵΪ "K3ListServer.clsLitUpdate.CopyBill.HookInvoke"
            'ErrDescription         | �����������̶�ֵ��ֵΪ "HookInvoke Error"
            'BillStatus             | ϵͳ���ã�����״̬�� 0 - δ���ˣ�1 - �ѹ���
            'ICClosed               | ϵͳ���ã�ҵ��ϵͳ�Ƿ���ɳ�ʼ��
            'UpStockWhenSave        | ϵͳ���ã������¿��ơ�True - ������£�False - ��˸���
            'FFinishInitAP          | Ӧ��Ӧ���Ƿ����á�0 - δ����
            'BillDict               | ���ݶ���
            '--------------------------------------------------------

        Case 200003 '����˳�ǰ�¼�
            '--------------------------------------------------------
 
            'dctParams ֵ��         | ����
            '--------------------------------------------------------
            'sDsn                   | �����ַ���
            'InterID                | ���ݵ�����
            'TranType               | ���ݵ���������
            'OperateDelete          | Ĭ��ֵΪ0
            'ROB                    | �Ƿ��Ǻ쵥
            'OperateCode            | ������:   1------���
            '                       |           2------�����
            '                       |           4------����
            '                       |           8------ɾ��
            '                       |           16-----������
            '                       |           32-----����
            '--------------------------------------------------------

        Case 200002 '��˷�д�¼�
            '--------------------------------------------------------
 
            'dctParams ֵ��         | ����
            '--------------------------------------------------------
            'sDsn                   | �����ַ���
            'InterID                | ���ݵ�����
            'TranType               | ���ݵ���������
            'OperateDelete          | Ĭ��ֵΪ0
            'ROB                    | �Ƿ��Ǻ쵥
            'OperateCode            | ������:   1------���
            '                       |           2------�����
            '                       |           4------����
            '                       |           8------ɾ��
            '                       |           16-----������
            '                       |           32-----����
            '--------------------------------------------------------

        Case 200001 '���ǰ�¼�
            '--------------------------------------------------------
 
            'dctParams ֵ��         | ����
            '--------------------------------------------------------
            'sDsn                   | �����ַ���
            'InterID                | ���ݵ�����
            'TranType               | ���ݵ���������
            'OperateDelete          | Ĭ��ֵΪ0
            'ROB                    | �Ƿ��Ǻ쵥
            'OperateCode            | ������:   1------���
            '                       |           2------�����
            '                       |           4------����
            '                       |           8------ɾ��
            '                       |           16-----������
            '                       |           32-----����
            '--------------------------------------------------------
            
            '�����������Ӵ��� '���ǰ' ����

 
    Case Else
    End Select
 
End Function
Public Function EndSave(ByVal sDsn As String, vctSave As KFO.Vector, sReturnMsg As String) As Boolean
    '-----------------------------------------------------------------------
    '����                   | ����
    '-----------------------------------------------------------------------
    'sDsn ����              | �����ַ���
    'sReturnMsg ����        | ���صĴ�����Ϣ
    'vctSave                | ����ʱ��Ϣ����ֵ�򹹳�����
    '-----------------------------------------------------------------------
    
    '-----------------------------------------------------------------------
    'vctSave ֵ��           | ����
    '-----------------------------------------------------------------------
    'FInterID               | ���ݵ�����
    'FTransType             | ���ݵ���������
    'IsRedBill              | ��ǰ�����Ƿ��Ǻ쵥
    'BillChecked            | ��ǰ�����Ƿ����
    'CurrentUserRight       | ��ǰ�û��Ƿ���Ȩ��
    'ActiveCtlValueChanged  | ��ǰ�����Ƿ�ı�
    'BeforeCloseStockBill   | ��ǰ�����Ƿ�������ǰ�ִ浥��(�ݹ���⣬���۳���)
    'ICClosed               | ��ǰ�Ƿ����
    '-----------------------------------------------------------------------
    
    On Error GoTo Catch
    '�����������Ӵ��� EndSave ����
    
    
    '���������ɣ���Ҫ���� True
    EndSave = True
    Exit Function
    
Catch:
    '������ִ�����Ҫ���� False���Լ���������
    EndSave = False
    sReturnMsg = Err.Description
End Function
 
Public Function BeginSave(ByVal sDsn As String, vctSave As KFO.Vector, sReturnMsg As String) As Boolean
    '-----------------------------------------------------------------------
    '����                   | ����
    '-----------------------------------------------------------------------
    'sDsn ����              | �����ַ���
    'sReturnMsg ����        | ���صĴ�����Ϣ
    'vctSave                | ����ʱ��Ϣ����ֵ�򹹳�����
    '-----------------------------------------------------------------------
    
    '-----------------------------------------------------------------------
    'vctSave ֵ��           | ����
    '-----------------------------------------------------------------------
    'FInterID               | ���ݵ�����
    'FTransType             | ���ݵ���������
    'IsRedBill              | ��ǰ�����Ƿ��Ǻ쵥
    'BillChecked            | ��ǰ�����Ƿ����
    'CurrentUserRight       | ��ǰ�û��Ƿ���Ȩ��
    'ActiveCtlValueChanged  | ��ǰ�����Ƿ�ı�
    'BeforeCloseStockBill   | ��ǰ�����Ƿ�������ǰ�ִ浥��(�ݹ���⣬���۳���)
    'ICClosed               | ��ǰ�Ƿ����
    '-----------------------------------------------------------------------
    
    On Error GoTo Catch
    '�����������Ӵ��� BeginSave ����
    
    
    '���������ɣ���Ҫ���� True
    BeginSave = True
    Exit Function
    
Catch:
    '������ִ�����Ҫ���� False���Լ���������
    BeginSave = False
    sReturnMsg = Err.Description
End Function
 