using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWMerge;
using ExcelWizard.Models.EWRow;
using ExcelWizard.Models.EWTable;
using System.Collections.Generic;

namespace ExcelWizard.Models.EWSheet;

public class Sheet : ISheetBuilder
{
    /// <summary>
    /// Sheet name
    /// </summary>
    public string? SheetName { get; internal set; }

    /// <summary>
    /// Insert one or more Table(s) data into the Sheet.
    /// Each Table is consist of some Rows and Cells with more style options to configure easily
    /// </summary>
    public List<Table> SheetTables { get; internal set; } = new();

    /// <summary>
    /// Insert one or more Row(s) records into the Sheet.
    /// Each Row is consist of some Cells with more style options to configure easily
    /// </summary>
    public List<Row> SheetRows { get; internal set; } = new();

    /// <summary>
    /// Insert one or more Cell(s) items directly into the Sheet.
    /// All data can be inserted with this property, but using  Tables and Rows add more options to configure style and functionality of generated Sheet.
    /// </summary>
    public List<Cell> SheetCells { get; internal set; } = new();

    /// <summary>
    /// Sheet style options like Direction, TextAlign, ColumnsDefaultWith, RowsDefaultHeight and etc. Also Columns style can be configured here
    /// </summary>
    public SheetStyle SheetStyle { get; internal set; } = new();

    /// <summary>
    /// Merged Cells in the Sheet. If it is possible to write Merges on the Table or Row models, manage to do that instead of here
    /// </summary>
    public List<MergedCells> MergedCellsList { get; internal set; } = new();

    /// <summary>
    /// Will override the ExcelBuilder AreSheetsLockedByDefault value. Default will inherit
    /// </summary>
    public bool? IsSheetLocked { get; internal set; }

    /// <summary>
    /// Set Sheet protection level
    /// </summary>
    public ProtectionLevel SheetProtectionLevel { get; internal set; } = new();
}