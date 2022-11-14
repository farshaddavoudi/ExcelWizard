using ExcelWizard.Models.EWSheet;
using System;
using System.Collections.Generic;

namespace ExcelWizard.Models;

public class ExcelBuilder : IExcelBuilder, IExpectGeneratingExcelTypeExcelBuilder, IExpectSheetsExcelBuilder
    , IExpectStyleExcelBuilder, IExpectOtherPropsAndBuildExcelBuilder, IExpectBuildExcelBuilder
{
    private ExcelBuilder() { }

    private ExcelModel ExcelModel { get; set; } = new();
    private bool CanBuild { get; set; }

    /// <summary>
    /// Set generated file name
    /// </summary>
    /// <param name="fileName"> Generated file name </param>
    public static IExpectGeneratingExcelTypeExcelBuilder SetGeneratedFileName(string? fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            throw new ArgumentException("Generated file name cannot be empty");

        return new ExcelBuilder
        {
            ExcelModel = new ExcelModel
            {
                GeneratedFileName = fileName
            }
        };
    }

    public IExpectSheetsExcelBuilder CreateGridLayoutExcel()
    {
        //if (string.IsNullOrWhiteSpace(sheetName))
        //    throw new ArgumentException("Sheet name cannot be empty");

        //return new SheetBuilder
        //{
        //    Sheet = new Sheet
        //    {
        //        SheetName = sheetName
        //    }
        //};
    }

    public IExpectSheetsExcelBuilder CreateComplexLayoutExcel()
    {
        return this;
    }

    public IExpectStyleExcelBuilder SetSheet(Sheet sheet)
    {
        CanBuild = true;

        ExcelModel.Sheets.Add(sheet);

        return this;
    }

    public IExpectStyleExcelBuilder SetSheets(List<Sheet> sheets)
    {
        if (sheets.Count == 0)
            throw new InvalidOperationException("Excel Sheets cannot be an empty list");

        CanBuild = true;

        ExcelModel.Sheets.AddRange(sheets);

        return this;
    }

    public IExpectOtherPropsAndBuildExcelBuilder SetSheetsDefaultStyle(SheetsDefaultStyle sheetsDefaultStyle)
    {
        ExcelModel.SheetsDefaultStyle = sheetsDefaultStyle;

        return this;
    }

    public IExpectOtherPropsAndBuildExcelBuilder NoDefaultStyle()
    {
        return this;
    }

    public IExpectBuildExcelBuilder SetDefaultLockedStatus(bool isLockedByDefault)
    {
        ExcelModel.AreSheetsLockedByDefault = isLockedByDefault;

        return this;
    }

    public ExcelModel Build()
    {
        if (CanBuild is false)
            throw new InvalidOperationException("Cannot build Excel model because some necessary information are not provided");

        return ExcelModel;
    }
}