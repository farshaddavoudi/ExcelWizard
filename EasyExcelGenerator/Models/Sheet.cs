using System.Collections.Generic;

namespace EasyExcelGenerator.Models;

public class Sheet
{
    /// <summary>
    /// Sheet name. Default will be Sheet1, Sheet2, etc
    /// </summary>
    public string? SheetName { get; set; }

    /// <summary>
    /// Insert one or more Table(s) data into the Sheet.
    /// Each Table is consist of some Rows and Cells with more style options to configure easily
    /// </summary>
    public List<Table> Tables { get; set; } = new();

    /// <summary>
    /// Insert one or more Row(s) records into the Sheet.
    /// Each Row is consist of some Cells with more style options to configure easily
    /// </summary>
    public List<Row> Rows { get; set; } = new();

    /// <summary>
    /// Insert one or more Cell(s) items directly into the Sheet.
    /// All data can be inserted with this property, but using  Tables and Rows add more options to configure style and functionality of generated Sheet.
    /// </summary>
    public List<Cell> Cells { get; set; } = new();

    /// <summary>
    /// Sheet style options like Direction, TextAlign, ColumnsDefaultWith, RowsDefaultHeight and etc
    /// </summary>
    public SheetStyle SheetStyle { get; set; } = new();

    /// <summary>
    /// Sheet specific Columns style like the Column Width, TextAlign, IsHidden, IsLocked, etc
    /// </summary>
    public List<SheetColumnStyle> SheetColumnsStyle { get; set; } = new();

    /// <summary>
    /// Merged Cells in the Sheet e.g. new List { "L16:L18" } will merge starting from L16 Cell until L18 Cell (MergeStartCell:MergeEndCell)
    /// </summary>
    public List<string> MergedCells { get; set; } = new();

    /// <summary>
    /// Will override the ExcelFileModel SheetsDefaultIsLocked value
    /// </summary>
    public bool? IsLocked { get; set; }

    /// <summary>
    /// Set Sheet protection level
    /// </summary>
    public ProtectionLevels ProtectionLevels { get; set; } = new();
}