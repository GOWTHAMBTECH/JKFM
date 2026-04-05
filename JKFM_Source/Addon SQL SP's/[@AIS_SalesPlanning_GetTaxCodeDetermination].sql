Create Procedure [dbo].[@AIS_SalesPlanning_GetTaxCodeDetermination](@ItemCode Varchar(100),@CardCode Varchar(100))
As 
Begin 

 Select   T0.TaxCode   From TCD3 T0 Inner Join TCD2 T1 On T0.Tcd2Id =T1.AbsId 
Inner Join TCD2 T2 on T2.AbsId =T0.AbsId  Where    T1.KeyFld_1_V =(Select U_TCategory  from OITM Where ItemCode=@ItemCode)
and T1.KeyFld_2_V=(Select State1 from OCRD Where CardCode=@CardCode)
End
 
 -- Select State1,State2 ,* from OCRD
 
 --Select  T2.KeyFld_1_V, KeyFld_2_V,*    From TCD2 T2