using System;


namespace Utilities.Office.Excel
{    
    // Note - PaddingBlankValues in the RowCollection, CellCollection and ColumnCollection classes:

    // Unused rows and cells aren't stored in the underlying XML of an Excel file, with good reason. 
    // When these classes are populated from the XML, the gaps aren't filled in. 
    // This causes a logic gap when it comes to traversing these collections (either with foreach or for).
    // There is a method in each class called PadBlank[ClassSpecific]s. This fills in the gaps.
    // This method is called when getting the Collection.Count or Collection.GetEnumerator to ensure the collection is complete when required.

    public abstract class BaseExcel : BaseOffice
    {        
                
    }    
}
