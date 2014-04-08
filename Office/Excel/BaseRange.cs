using System;
using System.Collections.Generic;

namespace Utilities.Office.Excel
{
    internal enum RangeStyleProperty : int
    {
        VerticalAlignment,
        WrapText,
        BorderId,
        Fill,
        Font
    }

    // At the moment only Cell inherits from BaseRange. Row and Column should do as well, but need to think about how it should work.
    public abstract class BaseRange : BaseExcel
    {
        public Worksheet Worksheet { get; private set; }
        
        //protected BaseRange Parent { get; set; }
        //internal IEnumerable<Cell> ChildCells { get; set; }

        // Range format properties. These are saved to range format objects when the document is saved.
        // This approach is inefficient - it's storing all of these properties for each range, when in fact they should be referencing common CellFormat objects.
        // An alternative approach could be to reassign or create a new CellFormat when one of these is updated. These properties still need to exist for the sake
        // of a logical API, but they don't need to store the data, just provide a channel to where it is stored (i.e. Cell.CellFormat...).

        // WHEN ADDING NEW PROPERTIES - update CopyFormatProperties, UpdateFormatProperties, EqualsStyle        
        
        // alignment    
        public HorizontalAlignmentOptions HorizontalAlignment 
        { 
            get { return CellFormat.Alignment.Horizontal; } 
            set { UpdateHorizontalAlignment(value); } 
        }
        public VerticalAlignmentOptions VerticalAlignment 
        { 
            get { return CellFormat.Alignment.Vertical; } 
            set { UpdateVerticalAlignment(value); } 
        }
        public bool WrapText 
        { 
            get { return CellFormat.Alignment.WrapText; } 
            set { UpdateWrapText(value); } 
        }

        // border
        public Borders Borders { get { return CellFormat.Borders; } }
        // would be nice to have something like this, although this implementation doesn't work. 
        // Use the UpdateOuterBorderStyle in the meantime
        //public Border OuterBorder 
        //{ 
        //    set 
        //    { 
        //        CellFormat.Borders.Left = value;
        //        CellFormat.Borders.Right = value;
        //        CellFormat.Borders.Top = value;
        //        CellFormat.Borders.Bottom = value; 
        //    } 
        //}        

        // fill
        public Fill Fill { get { return CellFormat.Fill; } }
        // font
        public Font Font { get { return CellFormat.Font; } }
        // number format
        public NumberFormat NumberFormat { get { return CellFormat.NumberFormat; } }  // TO BE CODED

        private int _styleIndex = CellFormat.DefaultStyleIndex;
        internal int StyleIndex { get { return _styleIndex; } set { _styleIndex = value; } }    // should stay internal

        private CellFormat CellFormat { get { return new CellFormat(this); } }
        private Styles Styles { get { return Worksheet.Workbook.Styles; } }

        public BaseRange(Worksheet worksheet, int styleIndex)
        {
            Worksheet = worksheet;
            StyleIndex = styleIndex;
        }
        //public BaseRange(Worksheet worksheet, int styleIndex, BaseRange parent) 
        //    : this(worksheet, styleIndex)
        //{
        //    Parent = parent;
        //}
        //public BaseRange(Worksheet worksheet, int styleIndex, IEnumerable<Cell> childCells)
        //    : this (worksheet, styleIndex)
        //{
        //    ChildCells = childCells;
        //}

        /***********************************
         * PUBLIC METHODS
         ************************************/

        public void UpdateOuterBorderStyle(BorderStyle borderStyle)
        {
            // this is very inefficient. The design is to make sure the collections update themselves on updating, but this 
            // needs to make changes and then save them. Is fine for the time being as the Border objects aren't big.
            CellFormat.Borders.Left.BorderStyle = borderStyle;
            CellFormat.Borders.Right.BorderStyle = borderStyle;
            CellFormat.Borders.Top.BorderStyle = borderStyle;
            CellFormat.Borders.Bottom.BorderStyle = borderStyle;
        }

        /***********************************
         * INTERNAL METHODS
         ************************************/

        internal bool EqualsStyle(BaseRange other)
        {
            return (other.VerticalAlignment == VerticalAlignment &&
                other.WrapText == WrapText &&
                other.Borders.Equals(Borders) &&
                other.Fill.Equals(Fill) &&
                other.Font.Equals(Font));
        }


        internal static string GetRangeAddress(Worksheet worksheet, string localAddress)
        {
            return string.Format("{0}!{1}", worksheet.EscapedName, localAddress);
        }

        internal static string GetLocalAddress(int address1, int address2, bool fixedReference = false)
        {
            return BaseRange.GetLocalAddress(address1.ToString(), address2.ToString(), fixedReference);
        }
        internal static string GetLocalAddress(string address1, string address2, bool fixedReference = false)
        {
            return string.Format("{0}{1}:{0}{2}", (fixedReference ? "$" : ""), address1, address2);
        }

        internal static string GetRangeAddress(Cell cell1, Cell cell2, bool fixedReference = false)
        {
            return BaseRange.GetLocalAddress(cell1.Address, cell2.Address, fixedReference);
        }
        internal static string GetRangeAddress(Row row1, Row row2, bool fixedReference = false)
        {
            return BaseRange.GetLocalAddress(row1.Index, row2.Index, fixedReference);
        }
        internal static string GetRangeAddress(Column column1, Column column2, bool fixedReference = false)
        {
            return BaseRange.GetLocalAddress(column1.Name, column2.Name, fixedReference);
        }

        /***********************************
         * PROTECTED METHODS
         ************************************/
        

        /***********************************
         * PRIVATE METHODS
         ************************************/

        private void UpdateHorizontalAlignment(HorizontalAlignmentOptions horizontalAlignment)
        {
            CellFormat.Alignment.Horizontal = horizontalAlignment;

            //if (ChildCells != null)
            //{
            //    foreach (Cell cell in ChildCells)
            //        cell.HorizontalAlignment = horizontalAlignment;
            //}
        }
        private void UpdateVerticalAlignment(VerticalAlignmentOptions verticalAlignment)
        {
            CellFormat.Alignment.Vertical = verticalAlignment;

            //if (ChildCells != null)
            //{
            //    foreach (Cell cell in ChildCells)
            //        cell.VerticalAlignment = verticalAlignment;
            //}
        }
        private void UpdateWrapText(bool wrapText)
        {
            CellFormat.Alignment.WrapText = wrapText;

            //if (ChildCells != null)
            //{
            //    foreach (Cell cell in ChildCells)
            //        cell.WrapText = wrapText;
            //}
        }
    }
}
