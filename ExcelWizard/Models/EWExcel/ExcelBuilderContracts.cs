using ExcelWizard.Models.EWGridLayout;
using ExcelWizard.Models.EWSheet;
using System.Collections.Generic;

namespace ExcelWizard.Models;

public interface IExcelBuilder
{
}

public interface IExpectGeneratingExcelTypeExcelBuilder
{
    /// <summary>
    /// Generate simple Grid layout Excel file. You have the choice of using data binding
    /// </summary>
    IExpectGridLayoutExcelBuilder CreateGridLayoutExcel();

    /// <summary>
    /// The Excel is not a grid layout Excel, therefore cannot be created through binding to a model and would be created composing different sub-components and configs
    /// </summary>
    IExpectSheetsExcelBuilder CreateComplexLayoutExcel();
}

public interface IExpectSheetsExcelBuilder
{
    /// <summary>
    /// Add one or more sheets to the Excel. It is required as Excel definition cannot be without Sheet(s).
    /// </summary>
    /// <param name="sheetBuilder"> SheetBuilder with Build() method at the end </param>
    /// <param name="sheetBuilders"> SheetBuilder(s) with Build() method at the end of them </param>
    IExpectStyleExcelBuilder SetSheets(ISheetBuilder sheetBuilder, params ISheetBuilder[] sheetBuilders);
}

public interface IExpectStyleExcelBuilder
{
    /// <summary>
    /// All Sheets shared default styles including default ColumnWidth, default RowHeight and sheets language Direction
    /// </summary>
    IExpectOtherPropsAndBuildExcelBuilder SetSheetsDefaultStyle(SheetsDefaultStyle sheetsDefaultStyle);

    /// <summary>
    /// No custom default styles for sheet(s) will be set. Styles can be set on each Sheet individually
    /// </summary>
    IExpectOtherPropsAndBuildExcelBuilder SheetsHaveNoDefaultStyle();
}

public interface IExpectOtherPropsAndBuildExcelBuilder : IExpectBuildExcelBuilder
{
    /// <summary>
    /// Set the default IsLocked value for all Sheets. Default is false
    /// </summary>
    IExpectBuildExcelBuilder SetDefaultLockedStatus(bool isLockedByDefault);
}

public interface IExpectBuildExcelBuilder
{
    IExcelBuilder Build();
}

public interface IExpectGridLayoutExcelBuilder
{
    /// <summary>
    /// Generate grid layout Excel using data binding, meaning we have a model which is configured with [ExcelSheet] and [ExcelSheetColumn] attributes
    /// </summary>
    /// <returns></returns>
    IExpectDataBoundGridLayoutExcelBuilder WithDataBinding();

    /// <summary>
    /// Generate grid layout Excel in usual way by create Sheets manually step by step and without data binding
    /// </summary>
    /// <returns></returns>
    IExpectSheetsExcelBuilder WithoutDataBinding();
}

public interface IExpectDataBoundGridLayoutExcelBuilder
{
    /// <summary>
    /// Add a sheet to Excel easily by binding a model as Excel data as well as configure it via attributes used in the model. Can be invoked multiple times for each Sheet
    /// </summary>
    /// <typeparam name="T">Type of the model supposed to be bound</typeparam>
    /// <param name="boundData">List of records supposed to bound to Excel, e.g. Persons, Phones, etc</param>
    /// <param name="sheetName">Name of the bound Sheet. If not set, the Sheet name will be gotten from SheetName property of [ExcelSheet] attribute</param>
    /// <returns></returns>
    IExpectAndAnotherSheetOrStyleExcelBuilder AddBoundSheet<T>(List<T> boundData, string? sheetName = default);

    /// <summary>
    /// Generate Grid layout Excel having multiple Sheets from same or different model types configured options with [ExcelSheet] and [ExcelSheetColumn] attributes
    /// </summary>
    /// <param name="boundSheets"> List of data list. e.g. object list of Persons, Phones, Universities, etc which each will be mapped to a Sheet </param>
    IExpectStyleExcelBuilder AddBoundSheets(List<BoundSheet> boundSheets);
}

public interface IExpectAnotherBoundSheetExcelBuilder
{
    /// <summary>
    /// Add a sheet to Excel easily by binding a model as Excel data as well as configure it via attributes used in the model. Can be invoked multiple times for each Sheet
    /// </summary>
    /// <typeparam name="T">Type of the model supposed to be bound</typeparam>
    /// <param name="boundData">List of records supposed to bound to Excel, e.g. Persons, Phones, etc</param>
    /// <param name="sheetName">Name of the bound Sheet. If not set, the Sheet name will be gotten from SheetName property of [ExcelSheet] attribute</param>
    /// <returns></returns>
    IExpectAndAnotherSheetOrStyleExcelBuilder AddAnotherBoundSheet<T>(List<T> boundData, string? sheetName = default);
}

public interface IExpectAndAnotherSheetOrStyleExcelBuilder : IExpectAnotherBoundSheetExcelBuilder, IExpectStyleExcelBuilder, IExpectBuildExcelBuilder
{
}
