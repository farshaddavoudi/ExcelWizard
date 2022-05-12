using BlazorDownloadFile;
using ClosedXML.Excel;
using ClosedXML.Report.Utils;
using ExcelWizard.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;
using Border = ExcelWizard.Models.Border;
using Color = System.Drawing.Color;
using Table = ExcelWizard.Models.Table;

namespace ExcelWizard.Service;

internal class ExcelWizardService : IExcelWizardService
{
    private readonly IBlazorDownloadFileService _blazorDownloadFileService;

    #region Constructor Injection

    public ExcelWizardService(IBlazorDownloadFileService blazorDownloadFileService)
    {
        _blazorDownloadFileService = blazorDownloadFileService;
    }

    #endregion

    public GeneratedExcelFile GenerateCompoundExcel(CompoundExcelBuilder compoundExcelBuilder)
    {
        using var xlWorkbook = ClosedXmlEngine(compoundExcelBuilder);

        // Save
        using var stream = new MemoryStream();

        xlWorkbook.SaveAs(stream);

        var content = stream.ToArray();

        if (compoundExcelBuilder.GeneratedFileName.IsNullOrWhiteSpace())
            compoundExcelBuilder.GeneratedFileName = $"ExcelWizard_{DateTime.Now:yyyy-MM-dd HH-mm-ss}";

        return new GeneratedExcelFile { FileName = compoundExcelBuilder.GeneratedFileName, Content = content };
    }

    public string GenerateCompoundExcel(CompoundExcelBuilder compoundExcelBuilder, string savePath)
    {
        using var xlWorkbook = ClosedXmlEngine(compoundExcelBuilder);

        if (compoundExcelBuilder.GeneratedFileName.IsNullOrWhiteSpace())
            compoundExcelBuilder.GeneratedFileName = $"ExcelWizard_{DateTime.Now:yyyy-MM-dd HH-mm-ss}";

        var saveUrl = $"{savePath}\\{compoundExcelBuilder.GeneratedFileName}.xlsx";

        // Save
        xlWorkbook.SaveAs(saveUrl);

        return saveUrl;
    }

    public GeneratedExcelFile GenerateGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder)
    {
        var compoundExcelBuilder = ConvertEasyGridExcelBuilderToExcelWizardBuilder(multiSheetsGridLayoutExcelBuilder);

        return GenerateCompoundExcel(compoundExcelBuilder);
    }

    public GeneratedExcelFile GenerateGridLayoutExcel(object singleSheetDataList, string? generatedFileName)
    {
        var gridLayoutExcelBuilder = new GridLayoutExcelBuilder
        {
            GeneratedFileName = generatedFileName,

            Sheets = new List<GridExcelSheet> { new() { DataList = singleSheetDataList } }
        };

        var compoundExcelBuilder = ConvertEasyGridExcelBuilderToExcelWizardBuilder(gridLayoutExcelBuilder);

        return GenerateCompoundExcel(compoundExcelBuilder);
    }

    public string GenerateGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder, string savePath)
    {
        var compoundExcelBuilder = ConvertEasyGridExcelBuilderToExcelWizardBuilder(multiSheetsGridLayoutExcelBuilder);

        return GenerateCompoundExcel(compoundExcelBuilder, savePath);
    }

    public string GenerateGridLayoutExcel(object singleSheetDataList, string savePath, string generatedFileName)
    {
        var gridLayoutExcelBuilder = new GridLayoutExcelBuilder
        {
            GeneratedFileName = generatedFileName,

            Sheets = new List<GridExcelSheet> { new() { DataList = singleSheetDataList } }
        };

        var compoundExcelBuilder = ConvertEasyGridExcelBuilderToExcelWizardBuilder(gridLayoutExcelBuilder);

        return GenerateCompoundExcel(compoundExcelBuilder, savePath);
    }

    public async Task<DownloadFileResult> BlazorDownloadGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder)
    {
        var generatedFile = GenerateGridLayoutExcel(multiSheetsGridLayoutExcelBuilder);

        return await _blazorDownloadFileService.DownloadFile(generatedFile.FileName, generatedFile.Content, TimeSpan.FromMinutes(2), generatedFile.Content?.Length ?? 0, generatedFile.MimeType);
    }

    public async Task<DownloadFileResult> BlazorDownloadGridLayoutExcel(object singleSheetDataList, string? generatedFileName = null)
    {
        var generatedFile = GenerateGridLayoutExcel(singleSheetDataList, generatedFileName);

        return await _blazorDownloadFileService.DownloadFile(generatedFile.FileName, generatedFile.Content, TimeSpan.FromMinutes(2), generatedFile.Content?.Length ?? 0, generatedFile.MimeType);
    }

    public async Task<DownloadFileResult> BlazorDownloadCompoundExcel(CompoundExcelBuilder compoundExcelBuilder)
    {
        var generatedFile = GenerateCompoundExcel(compoundExcelBuilder);

        return await _blazorDownloadFileService.DownloadFile(generatedFile.FileName, generatedFile.Content, TimeSpan.FromMinutes(2), generatedFile.Content?.Length ?? 0, generatedFile.MimeType);
    }


    #region Private Methods

    private XLWorkbook ClosedXmlEngine(CompoundExcelBuilder compoundExcelBuilder)
    {
        //-------------------------------------------
        //  Create Workbook (integrated with using statement)
        //-------------------------------------------
        var xlWorkbook = new XLWorkbook
        {
            RightToLeft = compoundExcelBuilder.AllSheetsDefaultStyle.AllSheetsDefaultDirection == SheetDirection.RightToLeft,
            ColumnWidth = compoundExcelBuilder.AllSheetsDefaultStyle.AllSheetsDefaultColumnWidth,
            RowHeight = compoundExcelBuilder.AllSheetsDefaultStyle.AllSheetsDefaultRowHeight
        };

        // Check any sheet available
        if (compoundExcelBuilder.Sheets.Count == 0)
            throw new Exception("No sheet is available to create Excel workbook");

        // Check sheet names are unique
        var sheetNames = compoundExcelBuilder.Sheets
            .Where(s => s.SheetName.IsNullOrWhiteSpace() is false)
            .Select(s => s.SheetName)
            .ToList();

        var uniqueSheetNames = sheetNames.Distinct().ToList();

        if (sheetNames.Count != uniqueSheetNames.Count)
            throw new Exception("Sheet names should be unique");

        // Auto naming for sheets

        int i = 1;
        foreach (Sheet sheet in compoundExcelBuilder.Sheets)
        {
            if (sheet.SheetName.IsNullOrWhiteSpace())
            {
                var isNameOk = false;

                while (isNameOk is false)
                {
                    var possibleName = $"Sheet{i}";

                    isNameOk = compoundExcelBuilder.Sheets.Any(s => s.SheetName == possibleName) is false;

                    if (isNameOk)
                        sheet.SheetName = possibleName;

                    i++;
                }
            }
        }

        //-------------------------------------------
        //  Add Sheets one by one to ClosedXML Workbook instance
        //-------------------------------------------
        foreach (var sheet in compoundExcelBuilder.Sheets)
        {
            // Set name
            var xlSheet = xlWorkbook.Worksheets.Add(sheet.SheetName);

            // Set protection level
            var protection = xlSheet.Protect(sheet.SheetProtectionLevel.Password);

            var atLeastOneItemAdded = false;

            // Local function to add to flag
            XLSheetProtectionElements AddToFlag(XLSheetProtectionElements allowedElements, XLSheetProtectionElements toAdd)
            {
                atLeastOneItemAdded = true;

                return allowedElements | toAdd;
            }

            XLSheetProtectionElements allowedElements = XLSheetProtectionElements.None;

            if (sheet.SheetProtectionLevel.DeleteColumns && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.DeleteColumns);
            if (sheet.SheetProtectionLevel.EditObjects && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.EditObjects);
            if (sheet.SheetProtectionLevel.FormatCells && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.FormatCells);
            if (sheet.SheetProtectionLevel.FormatColumns && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.FormatColumns);
            if (sheet.SheetProtectionLevel.FormatRows && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.FormatRows);
            if (sheet.SheetProtectionLevel.InsertColumns && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.InsertColumns);
            if (sheet.SheetProtectionLevel.InsertHyperLinks && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.InsertHyperlinks);
            if (sheet.SheetProtectionLevel.InsertRows && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.InsertRows);
            if (sheet.SheetProtectionLevel.SelectLockedCells && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.SelectLockedCells);
            if (sheet.SheetProtectionLevel.DeleteRows && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.DeleteRows);
            if (sheet.SheetProtectionLevel.EditScenarios && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.EditScenarios);
            if (sheet.SheetProtectionLevel.SelectUnlockedCells && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.SelectUnlockedCells);
            if (sheet.SheetProtectionLevel.Sort && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.Sort);
            if (sheet.SheetProtectionLevel.UseAutoFilter && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.AutoFilter);
            if (sheet.SheetProtectionLevel.UsePivotTableReports && !sheet.SheetProtectionLevel.HardProtect)
                allowedElements = AddToFlag(allowedElements, XLSheetProtectionElements.PivotTables);

            if (atLeastOneItemAdded)
                protection.AllowedElements = allowedElements;
            else
                protection.AllowNone();

            // Set direction
            if (sheet.SheetStyle.SheetDirection is not null)
                xlSheet.RightToLeft = sheet.SheetStyle.SheetDirection.Value == SheetDirection.RightToLeft;

            // Set default column width
            if (sheet.SheetStyle.SheetDefaultColumnWidth is not null)
                xlSheet.ColumnWidth = (double)sheet.SheetStyle.SheetDefaultColumnWidth;

            // Set default row height
            if (sheet.SheetStyle.SheetDefaultRowHeight is not null)
                xlSheet.RowHeight = (double)sheet.SheetStyle.SheetDefaultRowHeight;

            // Set visibility
            xlSheet.Visibility = sheet.SheetStyle.Visibility switch
            {
                SheetVisibility.Hidden => XLWorksheetVisibility.Hidden,
                SheetVisibility.VeryHidden => XLWorksheetVisibility.VeryHidden,
                _ => XLWorksheetVisibility.Visible
            };

            // Set TextAlign
            var textAlign = sheet.SheetStyle.SheetDefaultTextAlign ?? compoundExcelBuilder.AllSheetsDefaultStyle.AllSheetsDefaultTextAlign;

            xlSheet.Columns().Style.Alignment.Horizontal = textAlign switch
            {
                TextAlign.Center => XLAlignmentHorizontalValues.Center,
                TextAlign.Right => XLAlignmentHorizontalValues.Right,
                TextAlign.Left => XLAlignmentHorizontalValues.Left,
                TextAlign.Justify => XLAlignmentHorizontalValues.Justify,
                _ => throw new ArgumentOutOfRangeException()
            };

            //-------------------------------------------
            //  Columns properties
            //-------------------------------------------
            foreach (var columnStyle in sheet.SheetColumnsStyle)
            {
                // Infer XLAlignment from "ColumnProp"
                var columnAlignmentHorizontalValue = columnStyle.ColumnTextAlign switch
                {
                    TextAlign.Center => XLAlignmentHorizontalValues.Center,
                    TextAlign.Justify => XLAlignmentHorizontalValues.Justify,
                    TextAlign.Left => XLAlignmentHorizontalValues.Left,
                    TextAlign.Right => XLAlignmentHorizontalValues.Right,
                    _ => throw new ArgumentOutOfRangeException()
                };

                if (columnStyle.ColumnWidth is not null)
                {
                    if (columnStyle.ColumnWidth.WidthCalculationType == ColumnWidthCalculationType.AdjustToContents)
                        xlSheet.Column(columnStyle.ColumnNo).AdjustToContents();

                    else
                        xlSheet.Column(columnStyle.ColumnNo).Width = (double)columnStyle.ColumnWidth.Width!;
                }

                if (columnStyle.AutoFit)
                    xlSheet.Column(columnStyle.ColumnNo).AdjustToContents();

                if (columnStyle.IsColumnHidden)
                    xlSheet.Column(columnStyle.ColumnNo).Hide();

                xlSheet.Column(columnStyle.ColumnNo).Style.Alignment
                    .SetHorizontal(columnAlignmentHorizontalValue);
            }

            //-------------------------------------------
            //  Map Tables
            //-------------------------------------------
            foreach (var table in sheet.SheetTables)
            {
                foreach (var tableRow in table.TableRows)
                {
                    ConfigureRow(xlSheet, tableRow, sheet.SheetColumnsStyle, sheet.IsSheetLocked ?? compoundExcelBuilder.AreSheetsLockedByDefault);
                }

                var tableRange = xlSheet.Range(table.StartCellLocation.RowNumber,
                    table.StartCellLocation.ColumnNumber,
                    table.EndLocation.RowNumber,
                    table.EndLocation.ColumnNumber);

                // Config Outside-Border
                XLBorderStyleValues? outsideBorder = GetXlBorderLineStyle(table.OutsideBorder.BorderLineStyle);

                if (outsideBorder is not null)
                {
                    tableRange.Style.Border.SetOutsideBorder((XLBorderStyleValues)outsideBorder);
                    tableRange.Style.Border.SetOutsideBorderColor(XLColor.FromColor(table.OutsideBorder.BorderColor));
                }

                // Config Inside-Border
                XLBorderStyleValues? insideBorder = GetXlBorderLineStyle(table.InlineBorder.BorderLineStyle);

                if (insideBorder is not null)
                {
                    tableRange.Style.Border.SetInsideBorder((XLBorderStyleValues)insideBorder);
                    tableRange.Style.Border.SetInsideBorderColor(XLColor.FromColor(table.InlineBorder.BorderColor));
                }

                // Apply table merges here
                foreach (var mergedCells in table.MergedCells)
                {
                    xlSheet.Range(mergedCells).Merge();
                }
            }

            //-------------------------------------------
            //  Map Rows 
            //-------------------------------------------
            foreach (var sheetRow in sheet.SheetRows)
            {
                ConfigureRow(xlSheet, sheetRow, sheet.SheetColumnsStyle, sheet.IsSheetLocked ?? compoundExcelBuilder.AreSheetsLockedByDefault);
            }

            //-------------------------------------------
            //  Map Cells
            //-------------------------------------------
            foreach (var cell in sheet.SheetCells)
            {
                if (cell.IsCellVisible is false)
                    continue;

                ConfigureCell(xlSheet, cell, sheet.SheetColumnsStyle, sheet.IsSheetLocked ?? compoundExcelBuilder.AreSheetsLockedByDefault);
            }

            // Apply sheet merges here
            foreach (var mergedCells in sheet.MergedCells)
            {
                var rangeToMerge = xlSheet.Range(mergedCells).Cells();

                var value = rangeToMerge.FirstOrDefault(r => r.IsEmpty() is false)?.Value;

                rangeToMerge.First().SetValue(value);

                xlSheet.Range(mergedCells).Merge();
            }
        }

        return xlWorkbook;
    }

    private CompoundExcelBuilder ConvertEasyGridExcelBuilderToExcelWizardBuilder(GridLayoutExcelBuilder gridLayoutExcelBuilder)
    {
        var excelWizardBuilder = new CompoundExcelBuilder();

        if (gridLayoutExcelBuilder.GeneratedFileName.IsNullOrWhiteSpace() is false)
            excelWizardBuilder.GeneratedFileName = gridLayoutExcelBuilder.GeneratedFileName;

        foreach (var gridExcelSheet in gridLayoutExcelBuilder.Sheets)
        {
            if (gridExcelSheet.DataList is IEnumerable records)
            {
                var headerRow = new Row();

                var dataRows = new List<Row>();

                // Get Header 

                bool headerCalculated = false;

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

                    var excelWizardSheetAttribute = record.GetType().GetCustomAttribute<ExcelSheetAttribute>();

                    sheetName = excelWizardSheetAttribute?.SheetName;

                    sheetDirection = excelWizardSheetAttribute?.SheetDirection ?? SheetDirection.LeftToRight;

                    var defaultFontWeight = excelWizardSheetAttribute?.FontWeight;

                    var defaultFont = new TextFont
                    {
                        FontName = excelWizardSheetAttribute?.FontName,
                        FontSize = excelWizardSheetAttribute?.FontSize == 0 ? null : excelWizardSheetAttribute?.FontSize,
                        FontColor = Color.FromKnownColor(excelWizardSheetAttribute?.FontColor ?? KnownColor.Black),
                        IsBold = defaultFontWeight == FontWeight.Bold
                    };

                    isSheetLocked = excelWizardSheetAttribute?.IsSheetLocked ?? false;

                    var isSheetHardProtected = excelWizardSheetAttribute?.IsSheetHardProtected ?? false;

                    if (isSheetHardProtected)
                        sheetProtectionLevel.HardProtect = true;

                    borderType = excelWizardSheetAttribute?.BorderType ?? LineStyle.Thin;

                    var defaultTextAlign = excelWizardSheetAttribute?.DefaultTextAlign ?? TextAlign.Center;

                    PropertyInfo[] properties = record.GetType().GetProperties();

                    int xLocation = 1;

                    var recordRow = new Row
                    {
                        RowStyle = new RowStyle
                        {
                            RowHeight = excelWizardSheetAttribute?.DataRowHeight == 0 ? null : excelWizardSheetAttribute?.DataRowHeight,
                            BackgroundColor = excelWizardSheetAttribute?.DataBackgroundColor != null ? Color.FromKnownColor(excelWizardSheetAttribute.DataBackgroundColor) : Color.Transparent
                        }
                    };

                    // Each loop is a Column
                    foreach (var prop in properties)
                    {
                        var excelWizardColumnAttribute = (ExcelColumnAttribute?)prop.GetCustomAttributes(true).FirstOrDefault(x => x is ExcelColumnAttribute);

                        if (excelWizardColumnAttribute?.Ignore ?? false)
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
                            FontName = excelWizardColumnAttribute?.FontName ?? defaultFont.FontName,
                            FontSize = excelWizardColumnAttribute?.FontSize is null || excelWizardColumnAttribute.FontSize == 0 ? defaultFont.FontSize : excelWizardColumnAttribute.FontSize,
                            FontColor = excelWizardColumnAttribute is null || excelWizardColumnAttribute.FontColor == KnownColor.Transparent
                            ? defaultFont.FontColor.Value
                            : Color.FromKnownColor(excelWizardColumnAttribute.FontColor),
                            IsBold = excelWizardColumnAttribute is null || excelWizardColumnAttribute.FontWeight == FontWeight.Inherit
                            ? defaultFont.IsBold
                            : excelWizardColumnAttribute.FontWeight == FontWeight.Bold
                        };

                        // Header
                        if (headerCalculated is false)
                        {
                            var headerFont = JsonSerializer.Deserialize<TextFont>(JsonSerializer.Serialize(finalFont));

                            headerFont.IsBold = excelWizardColumnAttribute is null || excelWizardColumnAttribute.FontWeight == FontWeight.Inherit
                                ? defaultFontWeight != FontWeight.Normal
                                : excelWizardColumnAttribute.FontWeight == FontWeight.Bold;

                            headerRow.Cells.Add(new Cell(xLocation, yLocation)
                            {
                                Value = excelWizardColumnAttribute?.HeaderName ?? prop.Name,
                                CellStyle = new CellStyle
                                {
                                    CellFont = headerFont,
                                    CellTextAlign = GetCellTextAlign(defaultTextAlign, excelWizardColumnAttribute?.HeaderTextAlign)
                                },
                                CellContentType = CellContentType.Text,
                            });

                            headerRow.RowStyle.RowHeight = excelWizardSheetAttribute?.HeaderHeight == 0 ? null : excelWizardSheetAttribute?.HeaderHeight;

                            headerRow.RowStyle.BackgroundColor = excelWizardSheetAttribute?.HeaderBackgroundColor != null ? Color.FromKnownColor(excelWizardSheetAttribute.HeaderBackgroundColor) : Color.Transparent;

                            headerRow.RowStyle.OutsideBorder = new Border { BorderColor = Color.Black, BorderLineStyle = borderType };

                            headerRow.RowStyle.InsideBorder = new Border { BorderColor = Color.Black, BorderLineStyle = borderType };

                            // Calculate Columns style
                            columnsStyle.Add(new ColumnStyle
                            {
                                ColumnNo = xLocation,
                                ColumnWidth = new ColumnWidth
                                {
                                    Width = excelWizardColumnAttribute?.ColumnWidth == 0 ? null : excelWizardColumnAttribute?.ColumnWidth,
                                    WidthCalculationType = excelWizardColumnAttribute is null || excelWizardColumnAttribute.ColumnWidth == 0 ? ColumnWidthCalculationType.AdjustToContents : ColumnWidthCalculationType.ExplicitValue
                                }
                            });
                        }

                        // Data
                        recordRow.Cells.Add(new Cell(xLocation, yLocation + 1)
                        {
                            Value = prop.GetValue(record),
                            CellContentType = excelWizardColumnAttribute?.ExcelDataContentType ?? CellContentType.Text,
                            CellStyle = new CellStyle
                            {
                                CellFont = finalFont,
                                CellTextAlign = GetCellTextAlign(defaultTextAlign, excelWizardColumnAttribute?.DataTextAlign)
                            }
                        });

                        xLocation++;
                    }

                    dataRows.Add(recordRow);

                    yLocation++;

                    headerCalculated = true;
                }

                excelWizardBuilder.Sheets.Add(new Sheet
                {
                    SheetName = sheetName,

                    SheetStyle = new SheetStyle { SheetDirection = sheetDirection },

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
                            InlineBorder = new Border { BorderLineStyle = borderType },
                            OutsideBorder = new Border { BorderLineStyle = borderType }
                        }
                    },

                    // Columns
                    SheetColumnsStyle = columnsStyle
                });
            }
            else
            {
                throw new Exception("GridExcelSheet object should be IEnumerable");
            }
        }

        return excelWizardBuilder;
    }

    private void ConfigureCell(IXLWorksheet xlSheet, Cell cell, List<ColumnStyle> columnProps, bool isSheetLocked)
    {
        // Infer XLDataType and value from "cell" CellType
        XLDataType? xlDataType;
        object cellValue = cell.Value;
        switch (cell.CellContentType)
        {
            case CellContentType.Number:
                xlDataType = XLDataType.Number;
                break;

            case CellContentType.Percentage:
                xlDataType = XLDataType.Text;
                cellValue = $"{cellValue}%";
                break;

            case CellContentType.Currency:
                xlDataType = XLDataType.Number;
                if (IsNumber(cellValue) is false)
                    throw new Exception("Cell with Currency CellType should be Number type");
                cellValue = Convert.ToDecimal(cellValue).ToString("##,###");
                break;

            case CellContentType.MiladiDate:
                xlDataType = XLDataType.DateTime;
                if (cellValue is not DateTime)
                    throw new Exception("Cell with MiladiDate CellType should be DateTime type");
                break;

            case CellContentType.Text:
            case CellContentType.Formula:
                xlDataType = XLDataType.Text;
                break;

            default: // = CellType.General
                xlDataType = null;
                break;
        }

        // Infer XLAlignment from "cell"
        XLAlignmentHorizontalValues? cellAlignmentHorizontalValue = cell.CellStyle.CellTextAlign switch
        {
            TextAlign.Center => XLAlignmentHorizontalValues.Center,
            TextAlign.Left => XLAlignmentHorizontalValues.Left,
            TextAlign.Right => XLAlignmentHorizontalValues.Right,
            TextAlign.Justify => XLAlignmentHorizontalValues.Justify,
            _ => null
        };

        // Get IsLocked property based on Sheet and Cell "IsLocked" prop
        bool? isLocked = cell.IsCellLocked;

        if (isLocked is null)
        { // Get from ColumnProps level
            var x = cell.CellLocation.ColumnNumber;

            var relatedColumnProp = columnProps.SingleOrDefault(c => c.ColumnNo == x);

            isLocked = relatedColumnProp?.IsColumnLocked;

            if (isLocked is null)
            { // Get from sheet level
                isLocked = isSheetLocked;
            }
        }

        //-------------------------------------------
        //  Map column per Cells loop cycle
        //-------------------------------------------
        var locationCell = xlSheet.Cell(cell.CellLocation.RowNumber, cell.CellLocation.ColumnNumber);

        if (xlDataType is not null)
            locationCell.SetDataType((XLDataType)xlDataType);

        if (cell.CellContentType == CellContentType.Formula)
            locationCell.SetFormulaA1(cellValue.ToString());
        else
            locationCell.SetValue(cellValue);

        locationCell.Style.Alignment.SetWrapText(cell.CellStyle.Wordwrap);

        locationCell.Style.Protection.Locked = (bool)isLocked;

        if (cellAlignmentHorizontalValue is not null)
            locationCell.Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)cellAlignmentHorizontalValue);

        // Set Vertical Middle Align
        locationCell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

        // Set Font
        if (cell.CellStyle.CellFont.FontColor is not null)
            locationCell.Style.Font.SetFontColor(XLColor.FromColor(cell.CellStyle.CellFont.FontColor.Value));

        if (cell.CellStyle.CellFont.FontSize is not null)
            locationCell.Style.Font.SetFontSize(cell.CellStyle.CellFont.FontSize.Value);

        if (cell.CellStyle.CellFont.IsBold is not null)
            locationCell.Style.Font.SetBold(cell.CellStyle.CellFont.IsBold.Value);

        if (cell.CellStyle.CellFont.FontName.IsNullOrWhiteSpace() is false)
            locationCell.Style.Font.SetFontName(cell.CellStyle.CellFont.FontName);
    }

    private void ConfigureRow(IXLWorksheet xlSheet, Row row, List<ColumnStyle> columnsStyleList, bool isSheetLocked)
    {
        foreach (var rowCell in row.Cells)
        {
            if (rowCell.IsCellVisible is false)
                continue;

            ConfigureCell(xlSheet, rowCell, columnsStyleList, isSheetLocked);
        }

        // Configure merged cells in the row
        foreach (var cellsToMerge in row.MergedCellsList)
        {
            // CellsToMerge example is "B2:D2"
            xlSheet.Range(cellsToMerge).Row(1).Merge();
        }

        if (row.Cells.Count != 0)
        {
            var xlRow = xlSheet.Row(row.Cells.First().CellLocation.RowNumber);
            if (row.RowStyle.RowHeight is not null)
                xlRow.Height = (double)row.RowStyle.RowHeight;

            var xlRowRange = xlSheet.Range(row.GetRowStartCellLocation().RowNumber,
                row.GetRowStartCellLocation().ColumnNumber,
                row.GetRowEndCellLocation().RowNumber,
                row.GetRowEndCellLocation().ColumnNumber);

            if (row.RowStyle.Font.FontColor is not null)
                xlRowRange.Style.Font.SetFontColor(XLColor.FromColor(row.RowStyle.Font.FontColor.Value));

            if (row.RowStyle.Font.FontSize is not null)
                xlRowRange.Style.Font.SetFontSize(row.RowStyle.Font.FontSize.Value);

            if (row.RowStyle.Font.IsBold is not null)
                xlRowRange.Style.Font.SetBold(row.RowStyle.Font.IsBold.Value);

            if (row.RowStyle.Font.FontName.IsNullOrWhiteSpace() is false)
                xlRowRange.Style.Font.SetFontName(row.RowStyle.Font.FontName);

            xlRowRange.Style.Fill.SetBackgroundColor(XLColor.FromColor(row.RowStyle.BackgroundColor));

            XLBorderStyleValues? outsideBorder = GetXlBorderLineStyle(row.RowStyle.OutsideBorder.BorderLineStyle);

            if (outsideBorder is not null)
            {
                xlRowRange.Style.Border.SetOutsideBorder((XLBorderStyleValues)outsideBorder);
                xlRowRange.Style.Border.SetOutsideBorderColor(XLColor.FromColor(row.RowStyle.OutsideBorder.BorderColor));
            }

            XLBorderStyleValues? insideBorder = GetXlBorderLineStyle(row.RowStyle.InsideBorder.BorderLineStyle);

            if (insideBorder is not null)
            {
                xlRowRange.Style.Border.SetInsideBorder((XLBorderStyleValues)insideBorder);
                xlRowRange.Style.Border.SetInsideBorderColor(XLColor.FromColor(row.RowStyle.InsideBorder.BorderColor));
            }


        }
    }

    private XLBorderStyleValues? GetXlBorderLineStyle(LineStyle borderLineStyle)
    {
        return borderLineStyle switch
        {
            LineStyle.DashDotDot => XLBorderStyleValues.DashDotDot,
            LineStyle.Thick => XLBorderStyleValues.Thick,
            LineStyle.Thin => XLBorderStyleValues.Thin,
            LineStyle.Dotted => XLBorderStyleValues.Dotted,
            LineStyle.Double => XLBorderStyleValues.Double,
            LineStyle.DashDot => XLBorderStyleValues.DashDot,
            LineStyle.Dashed => XLBorderStyleValues.Dashed,
            LineStyle.SlantDashDot => XLBorderStyleValues.SlantDashDot,
            LineStyle.None => XLBorderStyleValues.None,
            _ => null
        };
    }

    private bool IsNumber(object value)
    {
        return value is sbyte
               || value is byte
               || value is short
               || value is ushort
               || value is int
               || value is uint
               || value is long
               || value is ulong
               || value is float
               || value is double
               || value is decimal;
    }

    #endregion
}