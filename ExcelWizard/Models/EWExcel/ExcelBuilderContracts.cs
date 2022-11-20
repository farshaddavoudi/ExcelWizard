using ExcelWizard.Models.EWExcel;
using ExcelWizard.Models.EWSheet;
using System.Collections.Generic;

namespace ExcelWizard.Models;

public interface IExcelBuilder
{

}

public interface IExpectGeneratingExcelTypeExcelBuilder
{
    /// <summary>
    /// Generate simple Grid layout Excel file with the option of simply using a model bind and configure Excel with attributes
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
    /// Generate Simple Single Sheet Grid layout Excel file from special model configured options with [ExcelSheet] and [ExcelSheetColumn] attributes
    /// </summary>
    /// <param name="bindingListModel"> List of data (should be something like List{Person}()) </param>
    IExpectBuildExcelBuilder WithOneSheetUsingModelBinding(object bindingListModel);

    /// <summary>
    /// Generate Grid layout Excel having multiple Sheets from different model types configured options with [ExcelSheet] and [ExcelSheetColumn] attributes.
    /// Here Sheet name will be get from model [ExcelSheet] attribute.
    /// </summary>
    /// <param name="listOfBindingListModel"> List of data list. e.g. object list of Persons, Phones, Universities, etc which each will be mapped to a Sheet </param>
    IExpectStyleExcelBuilder WithMultipleSheetsUsingModelBinding(List<object> listOfBindingListModel);

    /// <summary>
    /// Generate Grid layout Excel having multiple Sheets from same model types configured options with [ExcelSheet] and [ExcelSheetColumn] attributes
    /// Here Sheet names can be get from argument and not from model attribute.
    /// </summary>
    /// <param name="bindingSheets"> List of data list. e.g. object list of Persons, Phones, Universities, etc which each will be mapped to a Sheet </param>
    IExpectStyleExcelBuilder WithMultipleSheetsUsingModelBinding(List<BindingSheet> bindingSheets);

    /// <summary>
    /// Generate Grid layout Excel in usual way by create Sheets manually step by step and without model binding
    /// </summary>
    /// <returns></returns>
    IExpectSheetsExcelBuilder ManuallyWithoutModelBinding();
}
