using System;

namespace Utilities.Office.Excel
{
    public enum HorizontalAlignmentOptions : int
    {
        General = 0,
        Left = 1,
        Center = 2,
        Right = 3,
        Fill = 4,
        Justify = 5,
        CenterContinuous = 6,
        Distributed = 7
    }

    public enum VerticalAlignmentOptions : int
    {
        Bottom = 0,
        Top = 1,
        Center = 2,
        Justify = 3,
        Distributed = 4
    }

    // Mapped to OpenXml.SheetStateValues
    public enum WorksheetVisibility : int
    {
        Visible = 0,
        Hidden = 1,
        VeryHidden = 2
    }

    internal enum BorderProperty : int
    {
        BorderStyle,
        Color
    }

    internal enum BorderType : int
    {
        Left,
        Right,
        Top,
        Bottom,
        Diagonal
    }

    // Mapped to OpenXmlSpreadsheet.BorderStyleValues
    public enum BorderStyle : int
    {
        None = 0,
        Thin = 1,
        Medium = 2,
        Dashed = 3,
        Dotted = 4,
        Thick = 5,
        Double = 6,
        Hair = 7,
        MediumDashed = 8,
        DashDot = 9,
        MediumDashDot = 10,
        DashDotDot = 11,
        MediumDashDotDot = 12,
        SlantDashDot = 13
    }

    internal enum FillProperty : int
    {
        PatternType,
        ForegroundColor,
        BackgroundColor
    }

    // Copied from OpenXml metadata
    public enum PatternType : int
    {
        None = 0,
        Solid = 1,
        MediumGray = 2,
        DarkGray = 3,
        LightGray = 4,
        DarkHorizontal = 5,
        DarkVertical = 6,
        DarkDown = 7,
        DarkUp = 8,
        DarkGrid = 9,
        DarkTrellis = 10,
        LightHorizontal = 11,
        LightVertical = 12,
        LightDown = 13,
        LightUp = 14,
        LightGrid = 15,
        LightTrellis = 16,
        Gray125 = 17,
        Gray0625 = 18
    }


    // Copied from OpenXml metadata, with addition of Blank
    public enum CellDataType : int
    {
        Blank = -1,
        Boolean = 0,
        Number = 1,
        Error = 2,
        SharedString = 3,
        String = 4,
        InlineString = 5,
        Date = 6
    }

    public enum ShapeAnchorType : int
    {
        AbsoluteAnchor,
        OneCellAnchor,
        TwoCellAnchor
    }

    internal struct Anchor
    {
        public int Col;
        public int ColOffset;
        public int Row;
        public int RowOffset;

        public Anchor(int col, int colOffset, int row, int rowOffset)
        {
            Col = col;
            ColOffset = colOffset;
            Row = row;
            RowOffset = rowOffset;
        }
    }

    public enum DefinedNameScope : int
    {
        Workbook,
        Worksheet
    }

    public enum WorksheetOrientation : int
    {
        None,
        Portrait,
        Landscape
    }

    public enum NumberFormatType : int
    {
        None = 0,
        General = 1,
        Text = 2,
        Numeric = 3,
        Fraction = 4,
        DateTime = 5
    }
}
