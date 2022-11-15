using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWMerge;
using ExcelWizard.Models.EWRow;
using ExcelWizard.Models.EWTable;

namespace ExcelWizard.Models.EWSheet;

public interface ISheetBuilder
{

}

public interface IExpectSetComponentsSheetBuilder
{
    /// <summary>
    /// Insert a Table into the Sheet.
    /// The Table is consist of some Rows and Cells with more style options
    /// </summary>
    IExpectSetComponentsSheetBuilder SetTable(ITableBuilder table);

    /// <summary>
    /// Insert one or more Table(s) data into the Sheet.
    /// Each Table is consist of some Rows and Cells with more style options to configure easily
    /// </summary>
    IExpectSetComponentsSheetBuilder SetTables(params ITableBuilder[] tables);

    /// <summary>
    /// Insert one Row record into the Sheet.
    /// The Row is consist of some Cells with more style options 
    /// </summary>
    IExpectSetComponentsSheetBuilder SetRow(IRowBuilder row);

    /// <summary>
    /// Insert one or more Row(s) records into the Sheet.
    /// Each Row is consist of some Cells with more style options to configure easily
    /// </summary>
    IExpectSetComponentsSheetBuilder SetRows(params IRowBuilder[] rows);

    /// <summary>
    /// Insert a Cell item directly into the Sheet.
    /// All data can be inserted with this property, but using  Tables and Rows add more options to configure style and functionality of generated Sheet.
    /// </summary>
    IExpectSetComponentsSheetBuilder SetCell(ICellBuilder cell);

    /// <summary>
    /// Insert one or more Cell(s) items directly into the Sheet.
    /// All data can be inserted with this property, but using  Tables and Rows add more options to configure style and functionality of generated Sheet.
    /// </summary>
    IExpectSetComponentsSheetBuilder SetCells(params ICellBuilder[] cells);

    /// <summary>
    /// Finish composing Sheet with smaller components i.e. Table(s), Row(s) and Cell(s)
    /// </summary>
    /// <returns></returns>
    IExpectStyleSheetBuilder NoMoreTablesRowsOrCells();
}

public interface IExpectStyleSheetBuilder
{
    /// <summary>
    /// Set Sheet style options like Direction, TextAlign, ColumnsDefaultWith, RowsDefaultHeight and etc. Also Columns style can be configured here
    /// </summary>
    IExpectOtherPropsAndBuildSheetBuilder SetStyle(SheetStyle sheetStyle);

    /// <summary>
    /// No custom styles for the Sheet, neither for the Sheet itself nor for its Columns
    /// </summary>
    IExpectOtherPropsAndBuildSheetBuilder NoCustomStyle();
}

public interface IExpectOtherPropsAndBuildSheetBuilder
{
    /// <summary>
    /// Will override the ExcelBuilder AreSheetsLockedByDefault value. Default will inherit
    /// </summary>
    IExpectOtherPropsAndBuildSheetBuilder SetLockedStatus(bool isLocked);

    /// <summary>
    /// Set Sheet protection level
    /// </summary>
    IExpectOtherPropsAndBuildSheetBuilder SetProtectionLevel(ProtectionLevel protectionLevel);

    /// <summary>
    /// Merged Cells in the Sheet.
    /// We prefer merging cells in Table or Row sub-models but in some scenarios this option would be helpful
    /// </summary>
    IExpectOtherPropsAndBuildSheetBuilder SetMergedCells(params IMergeBuilder[] merges);

    Sheet Build();
}