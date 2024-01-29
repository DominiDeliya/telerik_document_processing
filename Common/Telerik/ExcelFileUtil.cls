 
/*------------------------------------------------------------------------
   File        : ExcelFileUtil
   Purpose     : import / export files using telerik document processing library
   Syntax      : 
   Description : 
   Author(s)   : domin
   Created     : Fri Feb 10 15:14:48 CET 2023
   Notes       : 
 ----------------------------------------------------------------------*/


BLOCK-LEVEL ON ERROR UNDO, THROW.

USING OpenEdge.Core.Collections.IIterator FROM PROPATH.
USING OpenEdge.Core.Collections.StringCollection FROM PROPATH.
USING System.IO.FileMode FROM ASSEMBLY.
USING System.IO.FileStream FROM ASSEMBLY.
USING Telerik.Windows.Documents.Spreadsheet.Expressions.Functions.True FROM ASSEMBLY.
USING Telerik.Windows.Documents.Spreadsheet.FormatProviders.IWorkbookFormatProvider FROM ASSEMBLY.
USING Telerik.Windows.Documents.Spreadsheet.FormatProviders.OpenXml.Xlsx.XlsxFormatProvider FROM ASSEMBLY.
USING Telerik.Windows.Documents.Spreadsheet.Model.CellRange FROM ASSEMBLY.
USING Telerik.Windows.Documents.Spreadsheet.Model.ColumnSelection FROM ASSEMBLY.
USING Telerik.Windows.Documents.Spreadsheet.Model.RowSelection FROM ASSEMBLY.
USING Telerik.Windows.Documents.Spreadsheet.Model.Workbook FROM ASSEMBLY.
USING Telerik.Windows.Documents.Spreadsheet.Model.Worksheet FROM ASSEMBLY.

CLASS Common.Telerik.ExcelFileUtil: 

    /*------------------------------------------------------------------------------
     Purpose: export and import excel files using telerik document processing library
     Notes:
    ------------------------------------------------------------------------------*/
    DEFINE PUBLIC PROPERTY fileName AS CHARACTER NO-UNDO
        PUBLIC GET.
        PRIVATE SET.
    
    DEFINE PRIVATE PROPERTY Workbook   AS Workbook                NO-UNDO
        PRIVATE GET.
        PRIVATE SET.
    
    DEFINE PRIVATE PROPERTY Worksheet AS Worksheet NO-UNDO
        GET.
        PRIVATE SET.
        
    DEFINE PRIVATE PROPERTY SaveFormat AS IWorkbookFormatProvider NO-UNDO
        GET.
        PRIVATE SET.
        
    DEFINE PRIVATE PROPERTY FileStream AS FileStream      NO-UNDO
        GET.
        PRIVATE SET.

    DEFINE PRIVATE PROPERTY FileMod AS FileMode NO-UNDO
        GET.
        PRIVATE SET.
   
    CONSTRUCTOR PUBLIC ExcelFileUtil ( cFileName AS CHARACTER ):
        SUPER ().
        THIS-OBJECT:fileName = cFileName. 
        
        RETURN.
    END CONSTRUCTOR.

    /*------------------------------------------------------------------------------
     Purpose:add or get work book and work sheet
     Notes: 
    ------------------------------------------------------------------------------*/

    METHOD PUBLIC VOID AddWorkBook( TransType AS CHARACTER ):
        
        IF TransType = "Import" THEN
        DO:
            FileMod   = FileMode:Open.
            FileStream = NEW FileStream(THIS-OBJECT:fileName, FileMod).
            SaveFormat = NEW XlsxFormatProvider(). 
            Workbook   = SaveFormat:Import(FileStream).
            Worksheet  = Workbook:ActiveWorksheet.
            
        END.
        ELSE IF TransType = "Export" THEN
        DO:
            FILE-INFO:FILE-NAME = THIS-OBJECT:fileName.
            IF FILE-INFO:FULL-PATHNAME EQ ?
            THEN
            DO:
                Workbook   = NEW Workbook().
                Worksheet = Workbook:Worksheets:Add().
        
                SaveFormat = NEW XlsxFormatProvider().   
                FileMod   = FileMode:Create.
                FileStream = NEW FileStream(THIS-OBJECT:fileName, FileMod).
            END.
            ELSE
            DO:
                FileMod   = FileMode:OpenOrCreate.
                SaveFormat = NEW XlsxFormatProvider().   
                FileStream = NEW FileStream(THIS-OBJECT:fileName, FileMod).
                Workbook   = SaveFormat:Import(FileStream).
                Workbook:ActiveWorksheet = Workbook:Worksheets:GetByName("Sheet1").   
                Worksheet  = Workbook:ActiveWorksheet.
                SaveFormat = NEW XlsxFormatProvider().
            END.
            
        END.

        RETURN.
        
        CATCH e AS Progress.Lang.Error:
            MESSAGE "Error in AddWorkBook" + e:GetMessage(1)
            VIEW-AS ALERT-BOX.
        END CATCH.
    END METHOD. 
    
     

    /*------------------------------------------------------------------------------
     Purpose: save work book after all the operations
     Notes:
    ------------------------------------------------------------------------------*/

    METHOD PUBLIC VOID SaveWorkBook(TransType AS CHARACTER ):
        
        IF TransType = "Import" THEN
            SaveFormat:Import(FileStream).
        ELSE IF TransType = "Export" THEN
            SaveFormat:Export(Workbook,FileStream).
        RETURN.

        CATCH e AS Progress.Lang.Error:
            MESSAGE "Error in AddHeaders" + e:GetMessage(1)
            VIEW-AS ALERT-BOX.
        END CATCH.

    END METHOD.    

    /*------------------------------------------------------------------------------
     Purpose: add headers for newly created excel file
     Notes:
    ------------------------------------------------------------------------------*/

    METHOD PUBLIC VOID AddHeaders(Headers AS CHARACTER   ):
        
        DEFINE VARIABLE iCounter AS INTEGER NO-UNDO.
        DEFINE VARIABLE iColumns AS INTEGER NO-UNDO.
        
        iColumns = NUM-ENTRIES(Headers).
     
        DO iCounter = 1 TO iColumns:
        
            Worksheet:Cells:GetCellSelection(0, iCounter - 1):SetValueAsText(ENTRY(iCounter,Headers)).
            //Bold the headings
            Worksheet:Cells:GetCellSelection(0, iCounter - 1):SetIsBold(TRUE). 
        END.

        RETURN.
        
        CATCH e AS Progress.Lang.Error :
            MESSAGE e:GetMessage(1) + "Error in AddHeaders"
                VIEW-AS ALERT-BOX. 
        END CATCH.
        
    END METHOD. 
    
    /*------------------------------------------------------------------------------
     Purpose: get header details of imported excel file
     Notes:
    ------------------------------------------------------------------------------*/

    METHOD PUBLIC CHARACTER GetHeaders( ):
        
        DEFINE VARIABLE cHeaders AS CHARACTER NO-UNDO.
        DEFINE VARIABLE iCount   AS INTEGER   NO-UNDO.
        DEFINE VARIABLE cValue   AS CHARACTER NO-UNDO.
        
        iCount = 0.
        REPEAT :

            IF Worksheet:Cells:GetCellSelection(0, iCount):GetValue():VALUE:RawValue <> ""
                THEN
            DO:
                cValue = Worksheet:Cells:GetCellSelection(0, iCount):GetValue():VALUE:RawValue.
                cHeaders = cHeaders + "," + cValue.
                iCount = iCount + 1. 
                
            END.
            ELSE
                LEAVE.
                
        END.
                
        RETURN cHeaders.
        
        CATCH e AS Progress.Lang.Error :
                MESSAGE "Error in GetRowCount" + e:GetMessage(1)
                VIEW-AS ALERT-BOX.
        END CATCH.
        
    END METHOD. 

    /*------------------------------------------------------------------------------
     Purpose: add rows to the excel using string collection
     Notes:
    ------------------------------------------------------------------------------*/

    METHOD PUBLIC VOID AddRows( oStrCollection AS StringCollection ):
        
        DEFINE VARIABLE iCounter      AS INTEGER              NO-UNDO.
        DEFINE VARIABLE iColumns      AS INTEGER              NO-UNDO.
        DEFINE VARIABLE iRow          AS INTEGER              NO-UNDO.
        DEFINE VARIABLE itrCollection AS IIterator            NO-UNDO.
        DEFINE VARIABLE cStringLine   AS OpenEdge.Core.String NO-UNDO.
        DEFINE VARIABLE iRowCount     AS INTEGER              NO-UNDO.
        
        itrCollection     = oStrCollection:Iterator().
        iRowCount = Worksheet:UsedCellRange:RowCount.
        
        REPEAT WHILE itrCollection:HasNext():
            
            cStringLine = CAST(itrCollection:Next() , OpenEdge.Core.String).

            DO iCounter = 1 TO NUM-ENTRIES(cStringLine:ToString()):
                Worksheet:Cells:GetCellSelection(iRowCount, iCounter - 1):SetValue(ENTRY(iCounter,cStringLine:ToString())).
            END.

            iRowCount = iRowCount + 1.
  
        END.
        
        RETURN.
        CATCH e AS Progress.Lang.Error :
            MESSAGE "Error in GetRowCount" + e:GetMessage(1)
            VIEW-AS ALERT-BOX.
        END CATCH.
    END METHOD.    

    /*------------------------------------------------------------------------------
     Purpose: get row values 
     Notes:
    ------------------------------------------------------------------------------*/

    METHOD PUBLIC CHARACTER GetRows( RowIndex AS INTEGER,
                                     ColumnIndex AS INTEGER ):
        
        DEFINE VARIABLE cRowValue AS CHARACTER NO-UNDO.
        
        cRowValue = Worksheet:Cells:GetCellSelection(RowIndex, ColumnIndex):GetValue():VALUE:RawValue.
        
        RETURN cRowValue.
        CATCH e AS Progress.Lang.Error :
                MESSAGE "Error in GetRowCount" + e:GetMessage(1)
                VIEW-AS ALERT-BOX.
        END CATCH.
    END METHOD.

    /*------------------------------------------------------------------------------
     Purpose: get used row count 
     Notes:
    ------------------------------------------------------------------------------*/

    METHOD PUBLIC INTEGER GetRowCount(  ):
        
        DEFINE VARIABLE iRowCount AS INTEGER NO-UNDO.
        
        iRowCount = Worksheet:UsedCellRange:RowCount.
        
        RETURN iRowCount.
        CATCH e AS Progress.Lang.Error :
                MESSAGE "Error in GetRowCount" + e:GetMessage(1)
                VIEW-AS ALERT-BOX.
        END CATCH.

    END METHOD.

    /*------------------------------------------------------------------------------
     Purpose: delete specific column range
     Notes:
    ------------------------------------------------------------------------------*/

    METHOD PUBLIC VOID DeleteColumns( FromIndex AS INTEGER, 
                                      ToIndex AS INTEGER ):
                                            
        DEFINE VARIABLE colSelection AS ColumnSelection NO-UNDO.
        
        colSelection = Worksheet:Columns:GetColumnSelection(FromIndex,ToIndex).
  
        colSelection:Remove().
  
        RETURN.
        
        CATCH e AS Progress.Lang.Error :
                MESSAGE "Error in DeleteColumns" + e:GetMessage(1)
                VIEW-AS ALERT-BOX.
        END CATCH.
    END METHOD. 
    
    /*------------------------------------------------------------------------------
     Purpose: delete a specific column
     Notes:
    ------------------------------------------------------------------------------*/

    METHOD PUBLIC VOID DeleteColumns( ColNumber AS INTEGER ):
        
        DEFINE VARIABLE colSelection AS ColumnSelection NO-UNDO.
        
        colSelection = Worksheet:Columns:GetColumnSelection(ColNumber).
  
        colSelection:Remove().
  
        RETURN.
        
        CATCH e AS Progress.Lang.Error :
            MESSAGE "Error in DeleteColumns" + e:GetMessage(1)
            VIEW-AS ALERT-BOX.
        END CATCH.
        
    END METHOD. 
             
    /*------------------------------------------------------------------------------
     Purpose: delete selected range of rows
     Notes:
    ------------------------------------------------------------------------------*/

    METHOD PUBLIC VOID DeleteRows( FromIndex AS INTEGER, 
                                   ToIndex AS INTEGER ):
                                         
        DEFINE VARIABLE rowSelection AS RowSelection NO-UNDO.
        
        rowSelection = Worksheet:Rows:GetRowSelection(FromIndex,ToIndex).
  
        rowSelection:Remove().
  
        RETURN.
        
        CATCH e AS Progress.Lang.Error :
            MESSAGE "Error in DeleteRows" + e:GetMessage(1)
            VIEW-AS ALERT-BOX.
        END CATCH.
        
    END METHOD. 
    
    /*------------------------------------------------------------------------------
     Purpose: delete a specific row
     Notes:
    ------------------------------------------------------------------------------*/

    METHOD PUBLIC VOID DeleteRows( RowNumber AS INTEGER ):
                                         
        DEFINE VARIABLE rowSelection AS RowSelection NO-UNDO.
        
        rowSelection = Worksheet:Rows:GetRowSelection(RowNumber).
  
        rowSelection:Remove().
  
        RETURN.
        
        CATCH e AS Progress.Lang.Error :
            MESSAGE "Error in DeleteRows" + e:GetMessage(1)
            VIEW-AS ALERT-BOX.
        END CATCH.
        
    END METHOD. 
    
    /*------------------------------------------------------------------------------
     Purpose: add filter for specific column range
     Notes:
    ------------------------------------------------------------------------------*/

    METHOD PUBLIC VOID AddFilter( FromRowIndex AS INTEGER, 
                                  FromColumnIndex AS INTEGER, 
                                  ToRowIndex AS INTEGER, 
                                  ToColumnIndex AS INTEGER ):
        
        DEFINE VARIABLE filterRange AS CellRange NO-UNDO.
        
        filterRange = NEW CellRange(FromRowIndex,FromColumnIndex,
                                    ToRowIndex,ToColumnIndex ).

        Worksheet:Filter:FilterRange = filterRange.
        
        RETURN.
        
        CATCH e AS Progress.Lang.Error :
            MESSAGE "Error in AddFilter" + e:GetMessage(1)
            VIEW-AS ALERT-BOX.
        END CATCH.
        
    END METHOD.
    
    DESTRUCTOR PUBLIC ExcelFileUtil ( ):
        
    END DESTRUCTOR.
    
    FINALLY:
    
    END FINALLY.

END CLASS.