 create proc [@AIS_SalesPlanning_GetUnitPrice](@ItemCode Varchar(max),@CardCode Varchar(max),@SalesUom Varchar(max))
 as
 Begin
 Select T2.Price  from OUOM T0 inner join OITM T1 on T1.PriceUnit=T0.UomEntry --and T0.UomCode =T1.SalUnitMsr 
  inner join ITM1 T2 on T2.ItemCode =T1.ItemCode
 inner join OCRD T3 on T3.ListNum =T2.PriceList Where T1.ItemCode =@ItemCode and T3.CardCode =@CardCode and T0.UomCode=@SalesUom 
 union all
  Select T1.Price  from OUOM T0 inner join ITM9 T1 on T1.UomEntry =T0.UomEntry
  inner join OITM T4 on T4.ItemCode =T1.ItemCode  --and T0.UomCode =T4.SalUnitMsr 
  inner join ITM1 T2 on T2.ItemCode =T1.ItemCode and T2.PriceList =T1.PriceList 
 inner join OCRD T3 on T3.ListNum =T2.PriceList Where T1.ItemCode =@ItemCode and T3.CardCode =@CardCode and T0.UomCode=@SalesUom 
 end
 