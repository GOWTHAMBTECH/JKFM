 create Proc [@AIS_BillDelivery_GetSerialNumbers](@ItemCode Varchar(Max),@WhsCode Varchar(Max))
 As
 Begin
 select   t1.ItemCode,IntrSerial,Indate,Quantity,datediff(dd,Indate,GETDATE())'Days' from osri t0

INNER JOIN SRI1 T1

ON T0.SysSerial = T1.SysSerial

and T0.[ItemCode]=T1.[ItemCode]
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    
where Status='0' and T0.[ItemCode]=@ItemCode and T0.WhsCode =@WhsCode 
End