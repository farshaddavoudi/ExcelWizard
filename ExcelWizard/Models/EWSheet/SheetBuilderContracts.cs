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
    /// Insert one or more Table(s) data into the Sheet.
    /// Each Table is consist of some Rows and Cells with more style options to configure easily
    /// </summary>
    /// <param name="tableBuilders"> TableBuilder(s) with Build() method at the end of them </param>
    IExpectSetComponentsSheetBuilder SetTables(params ITableBuilder[] tableBuilders);

    /// <summary>
    /// Insert one or more Row(s) records into the Sheet.
    /// Each Row is consist of some Cells with more style options to configure easily
    /// </summary>
    /// <param name="rowBuilders"> RowBuilder(s) with Build() method at the end of them </param>
    IExpectSetComponentsSheetBuilder SetRows(params IRowBuilder[] rowBuilders);

    /// <summary>
    /// Insert one or more Cell(s) items directly into the Sheet.
    /// All data can be inserted with this property, but using  Tables and Rows add more options to configure style and functionality of generated Sheet.
    /// </summary>
    /// <param name="cellBuilders"> CellBuilder(s) with Build() method at the end of them </param>
    IExpectSetComponentsSheetBuilder SetCells(params ICellBuilder[] cellBuilders);

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
    IExpectOtherPropsAndBuildSheetBuilder SetSheetStyle(SheetStyle sheetStyle);

    /// <summary>
    /// No custom styles for the Sheet, neither for the Sheet itself nor for its Columns
    /// </summary>
    IExpectOtherPropsAndBuildSheetBuilder SheetHasNoCustomStyle();
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
    /// <param name="mergeBuilders"> MergeBuilder(s) with Build() method at the end of them </param>
    IExpectOtherPropsAndBuildSheetBuilder SetMergedCells(params IMergeBuilder[] mergeBuilders);

    ISheetBuilder Build();
}