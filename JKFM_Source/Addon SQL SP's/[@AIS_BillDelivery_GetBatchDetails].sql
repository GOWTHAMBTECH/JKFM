CREATE Procedure [dbo].[@AIS_BillDelivery_GetBatchDetails] ( @ItemCode Varchar(max),@BatchQty Decimal(18,4) 
,@WhsCode Varchar(max))  
AS
Begin
Create Table #TableTemp (BatchNum Varchar(max),BatchQty Decimal(18,4)  ) 
SET NOCOUNT ON;    
DECLARE @site_value Decimal(18,4)  
SET @site_value = @BatchQty
Declare @BatchNum varchar(max), @Quantity Decimal(18,4) 
DECLARE brand_cursor CURSOR FOR     
   SELECT T1.BatchNum  BatchNum 
,T1.Quantity -Sum(isnull(T2.BatchQty,0)) Quantity
FROM OIBT T1 left join BatchTempTab  T2 on T2.ItemCode=T1.ItemCode collate SQL_Latin1_General_CP1_CI_AS and  T2.WhsCode=T1.WhsCode collate SQL_Latin1_General_CP1_CI_AS
 and  T2.BatchNum=T1.BatchNum collate SQL_Latin1_General_CP1_CI_AS
 WHERE T1.ItemCode =  @ItemCode and T1.Quantity >0 and T1.WhsCode = @WhsCode 
 group by  T1.BatchNum ,T1.Quantity,T1.WhsCode
 having T1.Quantity -Sum(isnull(T2.BatchQty,0))   >0
ORDER BY T1.WhsCode ASC
OPEN brand_cursor  
FETCH NEXT FROM brand_cursor INTO @BatchNum ,@Quantity 
WHILE @@FETCH_STATUS = 0    
BEGIN    
If (@site_value <= @Quantity)
Begin
Insert Into #TableTemp Values ( @BatchNum ,@site_value  )
SET @site_value = 0
End 
Else  If (@site_value > @Quantity)
Begin
Insert Into #TableTemp Values ( @BatchNum ,@Quantity  )
SET @site_value = @site_value -@Quantity
End 
FETCH NEXT FROM brand_cursor  INTO  @BatchNum ,@Quantity      
end 
CLOSE brand_cursor;    
DEALLOCATE brand_cursor;    
insert into BatchTempTab (ItemCode,WhsCode,BatchNum,BatchQty)
Select @ItemCode,@WhsCode, BatchNum,BatchQty From #TableTemp Where BatchQty>0.0
Select   BatchNum,BatchQty From #TableTemp Where BatchQty>0.0
Drop table #TableTemp
End

----Create Table BatchTempTab (ItemCode Varchar(max),WhsCode Varchar(max),BatchNum Varchar(max),BatchQty Decimal(18,4)  ) 

--Select * from BatchTempTab
