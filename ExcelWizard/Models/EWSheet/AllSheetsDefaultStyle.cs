using ExcelWizard.Models.EWStyles;

namespace ExcelWizard.Models.EWSheet
{
    public class AllSheetsDefaultStyle
    {
        public SheetDirection AllSheetsDefaultDirection { get; set; } = SheetDirection.LeftToRight;

        public TextAlign AllSheetsDefaultTextAlign { get; set; } = TextAlign.Left;

        /// <summary>
        /// Default column width for the workbook.
        /// <para>All new worksheets will use this column width.</para>
        /// </summary>
        public double AllSheetsDefaultColumnWidth { get; set; } = 20;

        /// <summary>
        /// Default row height for the workbook.
        /// <para>All new worksheets will use this row height.</para>
        /// </summary>
        public double AllSheetsDefaultRowHeight { get; set; } = 15;
    }
}