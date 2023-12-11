
-----------Auto Mail CC Config UDT
----Get Branch
SELECT T0."BPLId", T0."BPLName" FROM OBPL T0 WHERE T0."Disabled" ='N'

----Get BP Name
Select T0."CardName" "BP Name" from OCRD T0 where T0."CardCode"=$[$3.U_BPCode.0]

----Get BP Code
Select T0."CardCode" "BP Code",T0."CardName" "BP Name" from OCRD T0 where T0."validFor"='Y' Order by T0."CardCode"

----Get BP Group Name
Select (Select "GroupName" from OCRG where "GroupCode"=T0."GroupCode") "GroupName" from OCRD T0 where T0."CardCode"=$[$3.U_BPCode.0]