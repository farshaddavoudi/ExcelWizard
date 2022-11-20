using ClosedXML.Report.Utils;
using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWColumn;
using ExcelWizard.Models.EWGridLayout;
using ExcelWizard.Models.EWRow;
using ExcelWizard.Models.EWSheet;
using ExcelWizard.Models.EWStyles;
using ExcelWizard.Models.EWTable;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text.Json;

namespace ExcelWizard.Models.EWExcel;

public class ExcelBuilder : IExpectGeneratingExcelTypeExcelBuilder, IExpectSheetsExcelBuilder
    , IExpectStyleExcelBuilder, IExpectOtherPropsAndBuildExcelBuilder, IExpectBuildExcelBuilder
    , IExpectGridLayoutExcelBuilder
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

    public IExpectGridLayoutExcelBuilder CreateGridLayoutExcel()
    {
        return this;
    }

    public IExpectBuildExcelBuilder WithOneSheetUsingModelBinding(object bindingListModel)
    {
        CanBuild = true;

        var gridLayoutExcelModel = new GridLayoutExcelModel
        {
            GeneratedFileName = ExcelModel.GeneratedFileName,

            Sheets = new List<GridExcelSheet> { new() { DataList = bindingListModel } }
        };

        ExcelModel = ConvertEasyGridExcelBuilderToExcelWizardBuilder(gridLayoutExcelModel);

        return this;
    }

    public IExpectStyleExcelBuilder WithMultipleSheetsUsingModelBinding(List<object> listOfBindingListModel)
    {
        CanBuild = true;

        List<GridExcelSheet> gridSheets = listOfBindingListModel.Select(l => new GridExcelSheet { DataList = l }).ToList();

        GridLayoutExcelModel gridLayoutExcelModel = new GridLayoutExcelModel
        {
            GeneratedFileName = ExcelModel.GeneratedFileName,

            Sheets = gridSheets
        };

        ExcelModel = ConvertEasyGridExcelBuilderToExcelWizardBuilder(gridLayoutExcelModel);

        return this;
    }

    public IExpectStyleExcelBuilder WithMultipleSheetsUsingModelBinding(List<BindingSheet> bindingSheets)
    {
        CanBuild = true;

        List<GridExcelSheet> gridSheets = bindingSheets.Select(bs => new GridExcelSheet
        { SheetName = bs.SheetName, DataList = bs.BindingListModel }).ToList();

        GridLayoutExcelModel gridLayoutExcelModel = new GridLayoutExcelModel
        {
            GeneratedFileName = ExcelModel.GeneratedFileName,

            Sheets = gridSheets
        };

        ExcelModel = ConvertEasyGridExcelBuilderToExcelWizardBuilder(gridLayoutExcelModel);

        return this;
    }

    public IExpectSheetsExcelBuilder ManuallyWithoutModelBinding()
    {
        return this;
    }

    public IExpectSheetsExcelBuilder CreateComplexLayoutExcel()
    {
        return this;
    }

    public IExpectStyleExcelBuilder SetSheets(ISheetBuilder sheetBuilder, params ISheetBuilder[] sheetBuilders)
    {
        ISheetBuilder[] sheets = new[] { sheetBuilder }.Concat(sheetBuilders).ToArray();

        CanBuild = true;

        ExcelModel.Sheets.AddRange(sheets.Select(s => (Sheet)s));

        return this;
    }

    public IExpectOtherPropsAndBuildExcelBuilder SetSheetsDefaultStyle(SheetsDefaultStyle sheetsDefaultStyle)
    {
        ExcelModel.SheetsDefaultStyle = sheetsDefaultStyle;

        return this;
    }

    public IExpectOtherPropsAndBuildExcelBuilder SheetsHaveNoDefaultStyle()
    {
        return this;
    }

    public IExpectBuildExcelBuilder SetDefaultLockedStatus(bool isLockedByDefault)
    {
        ExcelModel.AreSheetsLockedByDefault = isLockedByDefault;

        return this;
    }

    public IExcelBuilder Build()
    {
        if (CanBuild is false)
            throw new InvalidOperationException("Cannot build Excel model because some necessary information are not provided");

        return ExcelModel;
    }


    private ExcelModel ConvertEasyGridExcelBuilderToExcelWizardBuilder(GridLayoutExcelModel gridLayoutExcelModel)
    {
        var excelWizardBuilder = new ExcelModel();

        if (gridLayoutExcelModel.GeneratedFileName.IsNullOrWhiteSpace() is false)
            excelWizardBuilder.GeneratedFileName = gridLayoutExcelModel.GeneratedFileName;

        foreach (var gridExcelSheet in gridLayoutExcelModel.Sheets)
        {
            gridExcelSheet.ValidateGridExcelSheetInstance();

            if (gridExcelSheet.DataList is IEnumerable records)
            {
                var headerRow = new Row();

                var dataRows = new List<Row>();

                // Get Header 

                bool isHeaderAlreadyCalculated = false;

                int yLocation = 1;

                string? sheetName = null;

                var borderType = LineStyle.Thin;

                var columnsStyle = new List<ColumnStyle>();

                SheetDirection sheetDirection = SheetDirection.LeftToRight;

                bool isSheetLocked = false;

                ProtectionLevel sheetProtectionLevel = new();

                foreach (var record in records)
                {
                    // Each record is an entire row of Excel

                    var excelSheetAttribute = record.GetType().GetCustomAttribute<ExcelSheetAttribute>();

                    sheetName = string.IsNullOrWhiteSpace(gridExcelSheet.SheetName)
                        ? excelSheetAttribute?.SheetName
                        : gridExcelSheet.SheetName;

                    sheetDirection = excelSheetAttribute?.SheetDirection ?? SheetDirection.LeftToRight;

                    var defaultFontWeight = excelSheetAttribute?.FontWeight;

                    var defaultFont = new TextFont
                    {
                        FontName = excelSheetAttribute?.FontName,
                        FontSize = excelSheetAttribute?.FontSize == 0 ? null : excelSheetAttribute?.FontSize,
                        FontColor = Color.FromKnownColor(excelSheetAttribute?.FontColor ?? KnownColor.Black),
                        IsBold = defaultFontWeight == FontWeight.Bold
                    };

                    isSheetLocked = excelSheetAttribute?.IsSheetLocked ?? false;

                    var isSheetHardProtected = excelSheetAttribute?.IsSheetHardProtected ?? false;

                    if (isSheetHardProtected)
                        sheetProtectionLevel.HardProtect = true;

                    borderType = excelSheetAttribute?.BorderType ?? LineStyle.Thin;

                    var defaultTextAlign = excelSheetAttribute?.DefaultTextAlign ?? TextAlign.Center;

                    PropertyInfo[] properties = record.GetType().GetProperties();

                    int xLocation = 1;

                    var recordRow = new Row
                    {
                        RowStyle = new RowStyle
                        {
                            RowHeight = excelSheetAttribute?.DataRowHeight == 0 ? null : excelSheetAttribute?.DataRowHeight,
                            BackgroundColor = excelSheetAttribute?.DataBackgroundColor != null ? Color.FromKnownColor(excelSheetAttribute.DataBackgroundColor) : Color.Transparent
                        }
                    };

                    // Each loop is a Column
                    foreach (var prop in properties)
                    {
                        var excelSheetColumnAttribute = (ExcelSheetColumnAttribute?)prop.GetCustomAttributes(true).FirstOrDefault(x => x is ExcelSheetColumnAttribute);

                        if (excelSheetColumnAttribute?.Ignore ?? false)
                            continue;

                        TextAlign GetCellTextAlign(TextAlign defaultAlign, TextAlign? headerOrDataTextAlign)
                        {
                            return headerOrDataTextAlign switch
                            {
                                TextAlign.Inherit => defaultAlign,
                                _ => headerOrDataTextAlign ?? defaultAlign
                            };
                        }

                        var finalFont = new TextFont
                        {
                            FontName = excelSheetColumnAttribute?.FontName ?? defaultFont.FontName,
                            FontSize = excelSheetColumnAttribute?.FontSize is null || excelSheetColumnAttribute.FontSize == 0 ? defaultFont.FontSize : excelSheetColumnAttribute.FontSize,
                            FontColor = excelSheetColumnAttribute is null || excelSheetColumnAttribute.FontColor == KnownColor.Transparent
                            ? defaultFont.FontColor.Value
                            : Color.FromKnownColor(excelSheetColumnAttribute.FontColor),
                            IsBold = excelSheetColumnAttribute is null || excelSheetColumnAttribute.FontWeight == FontWeight.Inherit
                            ? defaultFont.IsBold
                            : excelSheetColumnAttribute.FontWeight == FontWeight.Bold
                        };

                        // Header
                        if (isHeaderAlreadyCalculated is false)
                        {
                            var headerFont = JsonSerializer.Deserialize<TextFont>(JsonSerializer.Serialize(finalFont));

                            headerFont.IsBold = excelSheetColumnAttribute is null || excelSheetColumnAttribute.FontWeight == FontWeight.Inherit
                                ? defaultFontWeight != FontWeight.Normal
                                : excelSheetColumnAttribute.FontWeight == FontWeight.Bold;

                            Cell headerCell = CellBuilder
                                .SetLocation(xLocation, yLocation)
                                .SetValue(excelSheetColumnAttribute?.HeaderName ?? prop.Name)
                                .SetCellStyle(new CellStyle
                                {
                                    Font = headerFont,
                                    CellTextAlign = GetCellTextAlign(defaultTextAlign,
                                        excelSheetColumnAttribute?.HeaderTextAlign)
                                })
                                .SetContentType(CellContentType.Text)
                                .Build();

                            headerRow.RowCells.Add(headerCell);

                            headerRow.RowStyle.RowHeight = excelSheetAttribute?.HeaderHeight == 0 ? null : excelSheetAttribute?.HeaderHeight;

                            headerRow.RowStyle.BackgroundColor = excelSheetAttribute?.HeaderBackgroundColor != null ? Color.FromKnownColor(excelSheetAttribute.HeaderBackgroundColor) : Color.Transparent;

                            headerRow.RowStyle.RowOutsideBorder = new Border { BorderColor = Color.Black, BorderLineStyle = borderType };

                            headerRow.RowStyle.InsideCellsBorder = new Border { BorderColor = Color.Black, BorderLineStyle = borderType };

                            // Calculate Columns style
                            columnsStyle.Add(new ColumnStyle(xLocation)
                            {
                                ColumnWidth = new ColumnWidth
                                {
                                    Width = excelSheetColumnAttribute?.ColumnWidth == 0 ? null : excelSheetColumnAttribute?.ColumnWidth,
                                    WidthCalculationType = excelSheetColumnAttribute is null || excelSheetColumnAttribute.ColumnWidth == 0 ? ColumnWidthCalculationType.AdjustToContents : ColumnWidthCalculationType.ExplicitValue
                                }
                            });
                        }

                        // Data
                        var dataCell = CellBuilder
                            .SetLocation(xLocation, yLocation + 1)
                            .SetValue(prop.GetValue(record))
                            .SetContentType(excelSheetColumnAttribute?.ExcelDataContentType ?? CellContentType.Text)
                            .SetCellStyle(new CellStyle
                            {
                                Font = finalFont,
                                CellTextAlign = GetCellTextAlign(defaultTextAlign,
                                    excelSheetColumnAttribute?.DataTextAlign)
                            })
                            .Build();

                        recordRow.RowCells.Add(dataCell);

                        xLocation++;
                    }

                    dataRows.Add(recordRow);

                    yLocation++;

                    isHeaderAlreadyCalculated = true;
                }

                excelWizardBuilder.Sheets.Add(new Sheet
                {
                    SheetName = sheetName,

                    SheetStyle = new SheetStyle { SheetDirection = sheetDirection, ColumnsStyle = columnsStyle },

                    IsSheetLocked = isSheetLocked,

                    SheetProtectionLevel = sheetProtectionLevel,

                    // Header Row
                    SheetRows = new List<Row> { headerRow },

                    // Table Data
                    SheetTables = new List<Table>
                    {
                        new()
                        {
                            TableRows = dataRows,
                            TableStyle= new TableStyle
                            {
                                TableOutsideBorder = new Border { BorderLineStyle = borderType }
                            }
                        }
                    }
                });
            }
            else
            {
                throw new Exception("GridExcelSheet object should be IEnumerable");
            }
        }

        return excelWizardBuilder;
    }
}