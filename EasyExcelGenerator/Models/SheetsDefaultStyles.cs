namespace EasyExcelGenerator.Models
{
    public class SheetsDefaultStyles
    {
        public SheetDirection Direction { get; set; } = SheetDirection.RightToLeft;

        public TextAlign TextAlign { get; set; } = TextAlign.Right;

        /// <summary>
        /// Default column width for the workbook.
        /// <para>All new worksheets will use this column width.</para>
        /// </summary>
        public double ColumnsWidth { get; set; } = 20;

        /// <summary>
        /// Default row height for the workbook.
        /// <para>All new worksheets will use this row height.</para>
        /// </summary>
        public double RowsHeight { get; set; } = 15;
    }
}