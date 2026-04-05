 
 
CREATE Procedure [@AIS_BillDelivery_LoadPreSalesOrder](@PlanDate Varchar(max) ,@BPGrp Varchar(max) ) 
As
Begin
 Select T0.DocEntry ,T0.DocNum ,T0.U_CardCode ,T0.U_CardName ,T1.U_ItemCode ,U_ItemName ,T3.U_ItemCategory ,T1.U_DefUnit 
 , Sum(isnull(T2.U_Qty,0.0) ) As BagQuantity,Sum(isnull(T2.U_Qty,0.0)) As TonQuantity
 ,T1.U_UnitPrice,T2.U_WhsCode ,T2.U_WhsName ,T1.U_TaxCode 
  from [@AIS_OPRE] T0 inner join [@AIS_PRE1] T1 on T1.DocEntry=T0.DocEntry 
 inner join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry  and T2.U_UniqID =T1.LineId 
 inner join OITM T3 on T3.ItemCode =T1.U_ItemCode 
 inner join OCRD T4 on T4.CardCode =T0.U_CardCode
 Where Convert(Varchar(10), T2.U_PlanDate,112) =Convert(Varchar(10),@PlanDate,112)and (T4.GroupCode =@BPGrp or @BPGrp ='-1')
  and isnull(T1.U_ItemCode ,'')!='' and isnull(T2.U_WhsCode ,'')!='' and isnull(T2.U_ReApprove,'N')='Y' and ISNULL (T2.U_Qty ,0.0)>0.0
    and Convert(Varchar(Max),T0.DocEntry)+'_'+T1.U_ItemCode+'_'+Convert(Varchar(10), T2.U_PlanDate,112) not in (
  Select A1.U_PreSaleEnt+'_'+A1.U_ItemCode+'_'+Convert(Varchar(10), A0.U_BillRecDate,112) from  [@AIS_LOAD] A0 inner join
   [@AIS_LOAD1] A1 on A1.DocEntry =A0.DocEntry  Where A1.U_PreSaleEnt =T0.DocEntry  and isnull(A1.U_Select,'N') ='Y')  
    and isnull(T1.U_MobStatus ,'N')!='D' and isnull(t2.U_CompStatus,'-1')='-1' and isnull(t2.U_MobStatus,'-1')='-1'
 group by  T0.DocEntry ,T0.DocNum ,T0.U_CardCode ,T0.U_CardName ,T1.U_ItemCode ,U_ItemName 
 ,T2.U_WhsCode ,T2.U_WhsName ,T1.U_TaxCode ,U_UnitPrice ,T3.U_ItemCategory,T1.U_DefUnit 
End
  
   