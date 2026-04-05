 

 Create Procedure [dbo].[@AIS_SalePlanning_LoadPreSalesOrder](@FromDate Varchar(max),@ToDate Varchar(max))

 As

 Begin

  Select  T1.LineId as LineID, T1.U_ItemCode,T2.OnHand ,isnull(T2.U_ItemCategory,'') as ItemCategory ,T1.U_ItemName ,T1.U_DefUnit as U_PlanUnit    ,T1.DocEntry,Convert(Varchar(10),T0.U_FromDate,112) as FromDate,

  Convert(Varchar(10),T0.U_ToDate ,112) as ToDate    from [@AIS_OPRE] T0 inner join [@AIS_PRE1] T1 on T1.DocEntry =T0.DocEntry  

  inner join OITM T2 on T2.ItemCode =T1.U_ItemCode

  Where ISNULL (T1.U_ItemCode ,'')!=''

  And  Convert(Varchar(10),T0.U_FromDate,111)<= @FromDate and  Convert(Varchar(10),T0.U_ToDate,112)>=@ToDate

 End
