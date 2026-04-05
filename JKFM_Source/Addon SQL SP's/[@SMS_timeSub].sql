CREATE FUNCTION  [dbo].[@SMS_timeSub](@InTime Decimal(10,2), @OutTime Decimal(10,2))    

returns Decimal(10,2) as    

BEGIN    

declare @InTime_Dec Decimal(10,2)    

declare @OutTime_Dec Decimal(10,2)    

    

declare @InTime_Frac Decimal(10,2)    

declare @OutTime_Frac Decimal(10,2)    

    

declare @TotalTime Decimal(10,2)    

declare @TotalTime_DicPart Decimal(10,2)      

declare @TotalTime_FraPart Decimal(10,2)     

set @InTime_Dec =CONVERT(int,@InTime)    

set @OutTime_Dec =CONVERT(int,@OutTime)    

set @InTime_Frac =@InTime - @InTime_Dec    

set @OutTime_Frac =@OutTime - @OutTime_Dec    

set @TotalTime_DicPart = @OutTime_Dec -@InTime_Dec    

set @TotalTime_FraPart =@OutTime_Frac-@InTime_Frac    

set @TotalTime  = @TotalTime_DicPart + case when @TotalTime_FraPart < 0 then 0.60 - (@TotalTime_FraPart*-1) else @TotalTime_FraPart end     

set @TotalTime  = @TotalTime - case when @TotalTime_FraPart < 0 then 1 else 0 end    

return( @TotalTime )    

END