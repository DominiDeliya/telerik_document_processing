
/*------------------------------------------------------------------------
    File        : exportExcel.p
    Purpose     : create excel file using telerik document processing library

    Syntax      :

    Description : 

    Author(s)   : domin
    Created     : Wed Sep 06 11:11:32 CEST 2023
    Notes       :
  ----------------------------------------------------------------------*/

/* ***************************  Definitions  ************************** */

BLOCK-LEVEL ON ERROR UNDO, THROW.
USING OpenEdge.Core.Collections.*.
/* ********************  Preprocessor Definitions  ******************** */


/* ***************************  Main Block  *************************** */      
DEFINE VARIABLE cHeaders       AS CHARACTER                    NO-UNDO.
DEFINE VARIABLE oExcelFileUtil AS Common.Telerik.ExcelFileUtil NO-UNDO.
DEFINE VARIABLE cFileName      AS CHARACTER                    NO-UNDO.
DEFINE VARIABLE oStrCollection AS StringCollection             NO-UNDO.

oStrCollection = NEW StringCollection().

FOR EACH Customer NO-LOCK:
    
    oStrCollection:Add(STRING(Customer.CustNum)
        + "," + Customer.Name 
        + "," + Customer.Address
        + "," + Customer.Address2 
        + "," + Customer.City 
        + "," + String(Customer.PostalCode) 
        + "," + Customer.Country
        + "," + Customer.EmailAddress 
        + "," + String(Customer.Phone)
        + "," + Customer.SalesRep 
        + "," + String(Customer.CreditLimit)).  
                            
END.

cHeaders = "CustNum,Name,Address,Address2,City,PostalCode,
            Country,Email Address , Phone, SalesRep, CreditLimit".
            
cFileName = "customer_details.xlsx".
       
oExcelFileUtil = NEW Common.Telerik.ExcelFileUtil(cFileName).
oExcelFileUtil:AddWorkBook("Export").
oExcelFileUtil:AddHeaders(cHeaders).
oExcelFileUtil:AddRows(oStrCollection).
oExcelFileUtil:AddFilter(0, 0, 5, 9).
oExcelFileUtil:SaveWorkBook("Export").

MESSAGE "Done..!"
    VIEW-AS ALERT-BOX.


    


