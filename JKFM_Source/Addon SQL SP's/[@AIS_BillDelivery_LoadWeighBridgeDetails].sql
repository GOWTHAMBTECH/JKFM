Create Procedure [dbo].[@AIS_BillDelivery_LoadWeighBridgeDetails](@PackingListNo Varchar(max),@TruckNo Varchar(max),@Unit Varchar(max),@WBType Varchar(max))
 As 
  Begin
  Declare @WBDatabase As nvarchar(max)
  if @WBType='172.16.10.22'
  Begin
  Set @WBDatabase='SALESHO\SQLEXPRESS2008'
  End
  if @WBType='172.16.20.2'
  Begin
  Set @WBDatabase ='SALESBO\SQLEXPRESS2008'
  End
  if @WBType='172.16.30.1'
  Begin
  Set @WBDatabase ='WAREHOUSE1\SQLEXPRESS2008'
  End
  Select T0.Weight1Indicator ,  T0.Weight1Date  As Weight1Date ,Sum(Cast(T0.Weight1 As decimal(18,5))) As Weight1,
  Sum(Cast(T0.Weight2 As decimal(18,5))) As Weight2,sum(Cast(T0.ItemWeight As decimal(18,5))) As ItemWeight,
    Weight2Date As Weight2Date,T0.Vehicle_ID ,T0.Registration ,UDFValue,'' as MobileNo,isnull(T0.DriverName,'') as DriverName
   from AIS_TRV4 T0  Where --T0.UDFID =1 and 
   T0.UDFValue =@PackingListNo and  T0.Registration=@TruckNo  and  T0.WeighBridge=@WBDatabase 
  Group by T0.Weight1Indicator ,Weight1Date,Weight2Date ,T0.Vehicle_ID ,T0.Registration,UDFValue 
  End