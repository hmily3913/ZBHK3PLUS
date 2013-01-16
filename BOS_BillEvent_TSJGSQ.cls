VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOS_BillEvent_TSJGSQ"
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

Private Sub m_BillInterface_Change(ByVal dct As KFO.IDictionary, ByVal dctFld As KFO.IDictionary, ByVal Col As Long, ByVal Row As Long, Cancel As Boolean)
 
    'TODO: 请在此处添加代码响应事件
    If Col = 4 And Row <> -1 And dctFld("FFLD") <> "" Then
        On Error GoTo Catch
        Dim strSql
        strSql = "EXECUTE sp_executesql N'Select top 1 t2.FInterID, t2.FEntryID,t2.FItemID, "
        strSql = strSql & " Round((Case When t2.FCuryID=1 Then t2.FPrice Else t2.FPrice/CAST(1 AS FLOAT) End),ti.FPriceDecimal) As FPrice,"
        strSql = strSql & " ( Case When tc.FOperator=''*'' Then Round(CAST(t4.FLowPrice AS FLOAT)*               ( CASE WHEN tc.FCurrencyID = 1 THEN 1                     ELSE ((Case When te.fexchangerate IS NOT NULL Then CAST(te.FExchangeRate AS FLOAT) Else CAST(tc.FExchangeRate AS FLOAT) End)/CAST(1 AS FLOAT)) END) * CAST(@dblCoefficient AS FLOAT),ti.FPriceDecimal)        Else Round(CAST(t4.FLowPrice AS FLOAT)/              ( CASE WHEN tc.FCurrencyID = 1 THEN 1                      ELSE ((Case When te.fexchangerate IS NOT NULL Then CAST(te.FExchangeRate AS FLOAT) Else CAST(tc.FExchangeRate AS FLOAT) End)/CAST(1 AS FLOAT)) END) * CAST(@dblCoefficient AS FLOAT),ti.FPriceDecimal) End) As FLowPrice,"
        strSql = strSql & " t4.FCanSell,t4.FLPriceCtrl,"
        strSql = strSql & " (1-sign(t2.FAuxPropID)) As FOrder4"
        strSql = strSql & "  From IcPrcPly t1"
        strSql = strSql & "  Inner Join IcPrcPlyEntry t2 On t1.FInterID=t2.FInterID"
        strSql = strSql & "  Inner Join t_ICItem ti On t2.FItemID=ti.FItemID"
        strSql = strSql & "  Inner Join t_MeasureUnit tm On t2.FUnitID=tm.FMeasureUnitID"
        strSql = strSql & "  Inner Join IcPrcOpt t3 On t1.FPlyType=t3.FKey And t3.FValue=''1'' Inner Join IcPrcPlyEntrySpec t4 On t2.FInterID=t4.FInterID And t2.FItemID=t4.FItemID And t2.FRelatedID=t4.FRelatedID"
        strSql = strSql & "  Inner Join t_Currency tc On t4.FLPriceCuryID=tc.FCurrencyID"
        strSql = strSql & "   Left Join t_ExchangeRateentry te On te.FBegDate<=convert(varchar(10),''" & Date & "'',120) and te.FEndDate >=convert(varchar(10),''" & Date & "'',120)"
        strSql = strSql & "                                    And te.FExchangeRateType= 1"
        strSql = strSql & "                                    And te.FCyTo=t4.FLPriceCuryID"
        strSql = strSql & "  Where ((t2.FRelatedID =0 AND t1.FSysTypeID IN(1,3))"
        strSql = strSql & "  OR (t2.FRelatedID =0 AND t1.FSysTypeID IN(501,30))"
        strSql = strSql & "  OR (t2.FRelatedID =0 AND t1.FSysTypeID = 1007735)"
        strSql = strSql & "  OR t1.FSysTypeID = -1000  ) And t2.FBegDate<=''" & Date & "'' And t2.FEndDate>=''" & Date & "'' and t2.FItemID=@lItemID  And t2.FAuxPropID in (0,@lAuxPropID)"
        strSql = strSql & "  And (( ((t2.FBegQty<=@dblQty And t2.FEndQty>=@dblQty ) Or (t2.FBegQty=0 And t2.FEndQty=0)))"
        strSql = strSql & "  Or  (t2.FUnitID=ti.FUnitID And ((t2.FBegQty<=@dblBaseQty And t2.FEndQty>=@dblBaseQty ) Or (t2.FBegQty=0 And t2.FEndQty=0))))"
        strSql = strSql & "   And (t2.FCuryID=1 Or t2.FCuryID=1) And t2.FChecked=1"
        strSql = strSql & " And FPeriodType = 0 "
        strSql = strSql & " Order by t1.FPri,t3.FSort,FOrder4,t2.FBegDate Desc'"
        strSql = strSql & " ,N'@lItemID INT,@lAuxPropID INT,@lUnitID INT ,@dblQty DECIMAL(28,10),@dblBaseQty DECIMAL(28,10),@dblCoefficient FLOAT'"
        strSql = strSql & " ,@lItemID=" & dctFld("FFLD") & ",@lAuxPropID=0,@lUnitID=0,@dblQty=0,@dblBaseQty=0,@dblCoefficient=1"
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        Set rs = m_BillInterface.K3Lib.GetData(strSql)
        If Not rs.EOF Then
            If rs("FPrice") = 0 Then
                MsgBox "当前文件价格为0，请检测！", vbCritical, m_BillInterface.K3Lib.LoadKDString("金蝶提示")
            Else
                m_BillInterface.SetFieldValue "FText3", rs("FPrice"), Row
            End If
        Else
            MsgBox "当前文件价格为0，请检测！", vbCritical, m_BillInterface.K3Lib.LoadKDString("金蝶提示")
        End If
    End If
    GoTo ExitSub
Catch:
    Set rs = Nothing
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, m_BillInterface.K3Lib.LoadKDString("金蝶提示")
    End If
ExitSub:
    Set rs = Nothing
End Sub
