CREATE Proc [@AIS_BillDelivery_TCSDetails](@CardCode Varchar(max),@TaxCode Varchar(max))
As
Begin
declare  @AccTransacion Decimal(18,5)=0
declare @TaxAmt Decimal(18,5)=0
declare @ExpTaxCode Varchar(max)=''
declare @FromDate Varchar(max)=''
declare @ToDate Varchar(max)=''
Select @FromDate =Convert(varchar(10),Min(F_RefDate),112),@ToDate =Convert(varchar(10),max(T_RefDate),112) from OFPR --Where Year(GetDate()) between F_RefDate and T_RefDate

SELECT @AccTransacion=SUM(T0."Debit") FROM JDT1 T0 WHERE T0."ShortName"=@CardCode  AND 
T0."RefDate">=@FromDate and T0."RefDate"<=@ToDate
  
  Select @TaxAmt=U_TCSPercent from OCRD T0 inner join [@AIS_TCS] T1 on T1.Code=T0.CmpPrivate 
  Where t0.CardCode=@CardCode 

 if exists( Select Code from OSTC T1 Where Rate =0 and Code=@TaxCode)
 Begin
  set @ExpTaxCode=@TaxCode
 end
 else
 Begin
 if exists (Select'a' from OSTC Where  @TaxCode like '%SGST%')
 begin
  Select @ExpTaxCode=Code from OSTC T1 Where Rate =0 and Code like '%EXEMPTED%'
 end
 if exists (Select'a' from OSTC Where  @TaxCode like '%IGST%')
 begin
  Select @ExpTaxCode=Code from OSTC T1 Where Rate =0 and Code like '%IGST%'
 end
 end

--Select @AccTransacion as TransValue ,@TaxAmt As TCSAmt ,@ExpTaxCode as TaxCode
Select 5000100 as TransValue ,@TaxAmt As TCSAmt ,@ExpTaxCode as TaxCode

end