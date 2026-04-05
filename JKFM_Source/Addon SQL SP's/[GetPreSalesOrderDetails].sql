 
Create Procedure [dbo].[GetPreSalesOrderDetails](@ItemCode NVARCHAR(MAX),@PlanDate DateTime,@Unit Varchar(max)
,@Category Varchar(max),@QtyType Varchar(max) )
AS
BEGIN
Declare @Warehouse1 Varchar(max)
Select @Warehouse1=U_WhsCode  from [@AIS_BRN1] Where U_BranchCode=@Unit
 
if @Category='I'
Begin
if @Unit='01'
Begin
if @QtyType='B'
Begin
Select Sum(isnull(T2.U_Qty,0.0)) As Qty,T0.DocNum ,T0.DocEntry ,T1.U_ItemCode ,CONVERT (Varchar(10), T2.U_PlanDate ,112)
As PlanDate,T0.U_CardCode ,T0.U_CardName ,isnull(T2.U_DefUnit,'') As DefUnit from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
T1.U_ItemCode =@ItemCode and  T2.U_WhsCode=@Warehouse1
   and  isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'

 Group by T0.DocNum ,T0.DocEntry ,T1.U_ItemCode , T2.U_PlanDate,T0.U_CardCode ,T0.U_CardName,isnull(T2.U_DefUnit,'')
 End
 Else if @QtyType='T'
Begin
Select  Sum(isnull(T2.U_Qty,0.0)*T3.SalPackUn) As Qty,T0.DocNum ,T0.DocEntry ,T1.U_ItemCode ,CONVERT (Varchar(10), T2.U_PlanDate ,112)
As PlanDate,T0.U_CardCode ,T0.U_CardName ,isnull(T2.U_DefUnit,'') As DefUnit from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
inner join OITM T3 on T3.ItemCode =T1.U_ItemCode 
Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
T1.U_ItemCode =@ItemCode and T2.U_WhsCode=@Warehouse1
    and isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'

 Group by T0.DocNum ,T0.DocEntry ,T1.U_ItemCode , T2.U_PlanDate,T0.U_CardCode ,T0.U_CardName,isnull(T2.U_DefUnit,'')
 End
END
else if @Unit='02'
Begin
if @QtyType='B'
Begin
Select Sum(isnull(T2.U_Qty,0.0)) As Qty,T0.DocNum ,T0.DocEntry ,T1.U_ItemCode ,CONVERT (Varchar(10), T2.U_PlanDate ,112)
As PlanDate,T0.U_CardCode ,T0.U_CardName ,isnull(T2.U_DefUnit,'') As DefUnit from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
T1.U_ItemCode =@ItemCode  and T2.U_WhsCode=@Warehouse1
   and  isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'

 Group by T0.DocNum ,T0.DocEntry ,T1.U_ItemCode , T2.U_PlanDate,T0.U_CardCode ,T0.U_CardName,isnull(T2.U_DefUnit,'')
 END
 Else if @QtyType='T'
Begin
Select  Sum(isnull(T2.U_Qty,0.0) *T3.SalPackUn ) As Qty,T0.DocNum ,T0.DocEntry ,T1.U_ItemCode ,CONVERT (Varchar(10), T2.U_PlanDate ,112)
As PlanDate,T0.U_CardCode ,T0.U_CardName ,isnull(T2.U_DefUnit,'') As DefUnit from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
inner join OITM T3 on  T3.ItemCode =T1.U_ItemCode 
Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
T1.U_ItemCode =@ItemCode  and T2.U_WhsCode=@Warehouse1
   and  isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'

 Group by T0.DocNum ,T0.DocEntry ,T1.U_ItemCode , T2.U_PlanDate,T0.U_CardCode ,T0.U_CardName,isnull(T2.U_DefUnit,'')
 END
END
End
Else if @Category='C'
Begin
if @Unit='01'
Begin
if @QtyType='B'
Begin
Select Sum(isnull(T2.U_Qty,0.0)) As Qty,T0.DocNum ,T0.DocEntry ,isnull(T3.U_ItemCategory,'') As U_ItemCode ,CONVERT (Varchar(10), T2.U_PlanDate ,112)
As PlanDate,T0.U_CardCode ,T0.U_CardName ,isnull(T2.U_DefUnit,'') As DefUnit from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
inner join OITM T3 on T3.ItemCode=T1.U_ItemCode 
Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
 T3.U_ItemCategory+'(1)'  =@ItemCode  and T2.U_WhsCode=@Warehouse1
   and  isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'

 Group by T0.DocNum ,T0.DocEntry ,T3.U_ItemCategory, T2.U_PlanDate,T0.U_CardCode ,T0.U_CardName,isnull(T2.U_DefUnit,'')
END
else if @QtyType='T'
Begin
Select Sum(isnull(T2.U_Qty,0.0)*T3.SalPackUn)  As Qty,T0.DocNum ,T0.DocEntry ,isnull(T3.U_ItemCategory,'') As U_ItemCode ,CONVERT (Varchar(10), T2.U_PlanDate ,112)
As PlanDate,T0.U_CardCode ,T0.U_CardName ,isnull(T2.U_DefUnit,'') As DefUnit from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
inner join OITM T3 on T3.ItemCode=T1.U_ItemCode 
Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
 T3.U_ItemCategory+'(1)'  =@ItemCode  and T2.U_WhsCode=@Warehouse1
   and  isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'

 Group by T0.DocNum ,T0.DocEntry ,T3.U_ItemCategory, T2.U_PlanDate,T0.U_CardCode ,T0.U_CardName,isnull(T2.U_DefUnit,'')
END
END
else if @Unit='02'
Begin
if @QtyType='B'
Begin
Select Sum(isnull(T2.U_Qty,0.0) ) As Qty,T0.DocNum ,T0.DocEntry ,isnull(T3.U_ItemCategory,'') As U_ItemCode 
,CONVERT (Varchar(10), T2.U_PlanDate ,112)
As PlanDate,T0.U_CardCode ,T0.U_CardName ,isnull(T2.U_DefUnit,'') As DefUnit from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
inner join OITM T3 on T3.ItemCode=T1.U_ItemCode 
Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
 T3.U_ItemCategory+'(1)'  =@ItemCode  and T2.U_WhsCode=@Warehouse1
   and  isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'

 Group by T0.DocNum ,T0.DocEntry ,T3.U_ItemCategory, T2.U_PlanDate,T0.U_CardCode ,T0.U_CardName,isnull(T2.U_DefUnit,'')
 END
 Else if @QtyType='T'
Begin
Select Sum(isnull(T2.U_Qty,0.0)*T3.SalPackUn ) As Qty,T0.DocNum ,T0.DocEntry ,isnull(T3.U_ItemCategory,'') As U_ItemCode 
,CONVERT (Varchar(10), T2.U_PlanDate ,112)
As PlanDate,T0.U_CardCode ,T0.U_CardName ,isnull(T2.U_DefUnit,'') As DefUnit from [@AIS_OPRE] T0 Inner Join [@AIS_PRE1] T1 On T0.DocEntry =T1.DocEntry 
Inner Join [@AIS_PRE2] T2 on T2.DocEntry =T1.DocEntry And T2.U_UniqID =T1.LineId 
inner join OITM T3 on T3.ItemCode=T1.U_ItemCode 
Where  CONVERT (Varchar(10), T2.U_PlanDate ,112)  =@PlanDate And 
 T3.U_ItemCategory+'(1)'  =@ItemCode  and T2.U_WhsCode=@Warehouse1
   and  isnull(T1.U_MobStatus ,'N')!='D'And   isnull(T2.U_Approval,'N')='Y'

 Group by T0.DocNum ,T0.DocEntry ,T3.U_ItemCategory, T2.U_PlanDate,T0.U_CardCode ,T0.U_CardName,isnull(T2.U_DefUnit,'')
 END
END
End
END