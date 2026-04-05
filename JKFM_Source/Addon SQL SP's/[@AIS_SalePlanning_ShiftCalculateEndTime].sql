 
--Exec [dbo].[@AIS_SWP_ShiftMaster_CalculateEndTime]'09:30','12:45','13:30',9,'f'
CREATE Procedure [dbo].[@AIS_SalePlanning_ShiftCalculateEndTime](@ShiftFromTime varchar(max),@BreakStartTime varchar(max),@BreakEndTime varchar(max),@Hours varchar(max),@LunchEnable Varchar(10))

As

Begin

Declare @BreakHrs Decimal(18,2),@ToTime varchar(max)
select @BreakHrs=dbo.[@SMS_timeSub](Replace(@BreakStartTime,':','.') ,Replace(@BreakEndTime,':','.'))
Declare @Hours_Actual Decimal(18,2)
IF (@LunchEnable='True')
Begin
SEt @Hours_Actual =Convert(Decimal(18,2) , @Hours )  +@BreakHrs
  

Select @ToTime=DATEADD(HH, Convert(int,@Hours_Actual), @ShiftFromTime)

Select CONVERT(VARCHAR(5),DATEADD(HH,-(0),@ToTime),108) as ToTime, DATEDIFF(HH , @ShiftFromTime, CONVERT(VARCHAR(5),DATEADD(HH,-(0),@ToTime),108)) as WorkHours

end 

else

Begin
SEt @Hours_Actual =Convert(Decimal(18,2) , @Hours ) 
Declare @Min as Decimal(18,2)
Set @Min=@BreakHrs  % 1 
select @Min= PARSENAME( CONVERT (Decimal(18,2), @BreakHrs )  % 1 , 1)
Select @ToTime=DATEADD(HOUR , Convert(int,@Hours_Actual), @ShiftFromTime)
Select CONVERT(VARCHAR(5),DATEADD(MINUTE  ,-(@Min),@ToTime),108) as ToTime, DATEDIFF(HH , @ShiftFromTime, CONVERT(VARCHAR(5),DATEADD(MINUTE,-(@Min),@ToTime),108)) as WorkHours
 
End 





 



  

End