using ExcelWizard.Models.EWSheet;
using System.Collections.Generic;

namespace ExcelWizard.Models;

public interface IExcelBuilder
{

}

public interface IExpectGeneratingExcelTypeExcelBuilder
{
    /// <summary>
    /// The Excel is not a grid layout Excel, therefore cannot be created through binding to a model and would be created composing different sub-components and configs
    /// </summary>
    /// <returns></returns>
    IExpectSheetsExcelBuilder CreateComplexLayoutExcel();
}

public interface IExpectSheetsExcelBuilder
{
    /// <summary>
    /// Add a Sheet to Excel
    /// </summary>
    IExpectStyleExcelBuilder SetSheet(Sheet sheet);

    /// <summary>
    /// Add Sheets to Excel
    /// </summary>
    IExpectStyleExcelBuilder SetSheets(List<Sheet> sheets);
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
    IExpectOtherPropsAndBuildExcelBuilder NoDefaultStyle();
}

public interface IExpectOtherPropsAndBuildExcelBuilder : IExpectBuildExcelBuilder
{
    IExpectBuildExcelBuilder SetDefaultLockedStatus(bool isLockedByDefault);
}

public interface IExpectBuildExcelBuilder
{
    ExcelModel Build();
}