CREATE  PROC [@AIS_BillDelivery_GetDocTotal](@DocEntry Varchar(max),@DocDueDate Varchar(max))
As
begin
Declare @PreSalesEntry Varchar(max),@Unit Varchar(max)
sELECT @PreSalesEntry=U_PreSalEnt ,@Unit=T0.U_Unit FROM ORDR T0
Where T0.DocNum=@DocEntry and T0.DocDueDate =@DocDueDate

Select SUM((T1.Quantity *T6.Price)+ ((T1.Quantity *T6.Price)*isnull(T4.Rate,0.0)/100)) from ORDR T0 inner join RDR1 T1 on T1.DocEntry =T0.DocEntry
inner join ITM1 T2 on T2.ItemCode =T1.ItemCode
inner join OCRD T3 on T3.CardCode =T0.CardCode and T3.ListNum =t2.PriceList 
 inner join ITM9 T6 on T6.ItemCode =T1.ItemCode AND t6.PriceList =t3.ListNum 
 inner join OITM T7 on T7.ItemCode =T1.ItemCode
 inner join OUOM T5 on T5.UomEntry =T6.UomEntry and T5.UomCode =T7.SalUnitMsr 
left join OSTC T4 on T4.Code =T1.TaxCode  
Where T0.U_PreSalEnt=@PreSalesEntry and T0.U_Unit =@Unit
--Where T0.DocNum=@DocEntry and T0.DocDueDate =@DocDueDate
end

--Exec dbo.[@AIS_BillDelivery_GetDocTotal]'202110174','20200918' 
--[@AIS_BillDelivery_GetDocTotal]'202110178',''

--Select SUM((T1.Quantity *T6.Price)+ ((T1.Quantity *T6.Price)*isnull(T4.Rate,0.0)/100)) from ORDR T0 inner join RDR1 T1 on T1.DocEntry =T0.DocEntry
--inner join ITM1 T2 on T2.ItemCode =T1.ItemCode
--inner join OCRD T3 on T3.CardCode =T0.CardCode and T3.ListNum =t2.PriceList 
-- inner join ITM9 T6 on T6.ItemCode =T1.ItemCode AND t6.PriceList =t3.ListNum 
-- inner join OITM T7 on T7.ItemCode =T1.ItemCode
-- inner join OUOM T5 on T5.UomEntry =T6.UomEntry and T5.UomCode =T7.SalUnitMsr 
--left join OSTC T4 on T4.Code =T1.TaxCode  Where T0.DocNum=202110178