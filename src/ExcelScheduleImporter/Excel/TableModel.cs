using System.Collections.Generic;
using DrawingColor = System.Drawing.Color;

namespace ExcelScheduleImporter.Excel
{
    /// <summary>Everything extracted from one Excel sheet region, in paper-space millimetres.</summary>
    public class TableModel
    {
        /// <summary>Width of each column in mm (printed size).</summary>
        public double[] ColWidthsMm { get; set; }

        /// <summary>Height of each row in mm (printed size).</summary>
        public double[] RowHeightsMm { get; set; }

        /// <summary>Master cells only (merged regions appear once, with spans).</summary>
        public List<CellData> Cells { get; set; } = new List<CellData>();

        /// <summary>
        /// Horizontal border segments. [r, c] = edge ABOVE row r at column c.
        /// Dimensions: [Rows + 1, Cols]. Weight 0 = no line.
        /// </summary>
        public int[,] HEdges { get; set; }

        /// <summary>
        /// Vertical border segments. [r, c] = edge LEFT of column c at row r.
        /// Dimensions: [Rows, Cols + 1]. Weight 0 = no line.
        /// </summary>
        public int[,] VEdges { get; set; }

        public int Rows => RowHeightsMm.Length;
        public int Cols => ColWidthsMm.Length;
        public string SourceSheet { get; set; }
        public string SourceRange { get; set; }
    }

    /// <summary>One visible cell (or merged region).</summary>
    public class CellData
    {
        // 0-based position within the extracted table
        public int Row { get; set; }
        public int Col { get; set; }
        public int RowSpan { get; set; } = 1;
        public int ColSpan { get; set; } = 1;

        public string Text { get; set; }
        public string FontName { get; set; } = "Arial";
        public double FontSizePt { get; set; } = 11.0;
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public DrawingColor TextColor { get; set; } = DrawingColor.Black;

        /// <summary>-1 = left, 0 = center, 1 = right</summary>
        public int HAlign { get; set; } = -1;

        /// <summary>-1 = top, 0 = middle, 1 = bottom</summary>
        public int VAlign { get; set; }

        /// <summary>Solid fill color, or null if no fill.</summary>
        public DrawingColor? Fill { get; set; }

        public bool WrapText { get; set; }
    }
}
