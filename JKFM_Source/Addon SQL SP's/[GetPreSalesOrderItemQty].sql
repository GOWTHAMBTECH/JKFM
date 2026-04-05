 
Create FUNCTION [dbo].[GetPreSalesOrderItemQty]
(     @ItemCode NVARCHAR(MAX)      	 
	 ,@PlanDate DateTime 
     , @Unit NVARCHAR(MAX),@Category  NVARCHAR(MAX)
	 ,@QtyType  NVARCHAR(MAX)	 	 
	  )
RETURNS @Output TABLE (  Qty Decimal(18,4) )
AS
BEGIN
Declare @Warehouse1 Varchar(max)
if @Unit='01' or @Unit='02'
Begin
Select @Warehouse1=U_WhsCode  from [@AIS_BRN1] Where U_BranchCode=@Unit
End
if @Category='I'
BEgin
if @Unit='01'
Begin
if @QtyType='B'
Begin
INSERT INTO @Output(Qty) 
 Select Sum(T2.U_Qty)  from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
 Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
 Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
     isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'and
  T1.U_ItemCode =@ItemCode and  T2.U_WhsCode =@Warehouse1 
End
Else if @QtyType='T'
Begin
INSERT INTO @Output(Qty) 
 Select Sum(T2.U_Qty*T3.SalPackUn)  from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
 Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
 inner join OITM T3 on T3.ItemCode =T1.U_ItemCode 
 Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
     isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'and

 T1.U_ItemCode =@ItemCode and  T2.U_WhsCode =@Warehouse1
End
End
 Else if @Unit='02'
Begin
if @QtyType='B'
Begin
INSERT INTO @Output( Qty) 
 Select Sum(T2.U_Qty)  from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
 Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
 Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
     isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'and

 T1.U_ItemCode =@ItemCode and  T2.U_WhsCode =@Warehouse1
End
Else if @QtyType='T'
Begin
INSERT INTO @Output( Qty) 
 Select Sum(T2.U_Qty*T3.SalPackUn )  from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
 Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
 inner join OITM T3 on T3.ItemCode =T1.U_ItemCode 
 Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
     isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'and

 T1.U_ItemCode =@ItemCode and  T2.U_WhsCode =@Warehouse1
End
End
End
else if @Category='C'
BEgin
if @Unit='01'
Begin
if @QtyType='B'
Begin
INSERT INTO @Output(Qty) 
 Select Sum(T2.U_Qty)  from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
 Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
 inner join OITM T3 on T3.ItemCode=T1.U_ItemCode 
 Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
     isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'and

 T3.U_ItemCategory+'(1)'   =@ItemCode and  T2.U_WhsCode =@Warehouse1
End
if @QtyType='T'
Begin
INSERT INTO @Output(Qty) 
 Select Sum(T2.U_Qty*T3.SalPackUn )  from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
 Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
 inner join OITM T3 on T3.ItemCode=T1.U_ItemCode 
 Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
     isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'and

 T3.U_ItemCategory+'(1)'   =@ItemCode and    T2.U_WhsCode =@Warehouse1
End
End
 Else if @Unit='02'
Begin
if @QtyType='B'
Begin
INSERT INTO @Output( Qty) 
 Select Sum(T2.U_Qty)  from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
 Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
 inner join OITM T3 on T3.ItemCode=T1.U_ItemCode 
 Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
     isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'and

 T3.U_ItemCategory+'(1)'   =@ItemCode and  T2.U_WhsCode =@Warehouse1
End
Else if @QtyType='T'
Begin
INSERT INTO @Output( Qty) 
 Select Sum(T2.U_Qty*T3.SalPackUn )  from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
 Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
 inner join OITM T3 on T3.ItemCode=T1.U_ItemCode 
 Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
     isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'and

 T3.U_ItemCategory+'(1)'   =@ItemCode and    T2.U_WhsCode =@Warehouse1
End
End
End
RETURN

END

  