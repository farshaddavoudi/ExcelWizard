using ClosedXML.Excel;
using ClosedXML.Report.Utils;
using EasyExcelGenerator.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using Border = EasyExcelGenerator.Models.Border;
using Color = System.Drawing.Color;
using Table = EasyExcelGenerator.Models.Table;

namespace EasyExcelGenerator.Service;

// TODO: Remove static and make them work with DI
public static class EasyExcelService
{
    /// <summary>
    /// Generate Excel file into file result
    /// </summary>
    /// <param name="easyExcelBuilder"></param>
    /// <returns></returns>
    public static GeneratedExcelFile GenerateExcel(EasyExcelBuilder easyExcelBuilder)
    {
        using var xlWorkbook = ClosedXmlEngine(easyExcelBuilder);

        // Save
        using var stream = new MemoryStream();

        xlWorkbook.SaveAs(stream);

        var content = stream.ToArray();

        return new GeneratedExcelFile { Content = content };
    }

    /// <summary>
    /// Generate Excel file and save it in path and return the saved url
    /// </summary>
    /// <param name="easyExcelBuilderFile"></param>
    /// <param name="savePath"></param>
    /// <returns></returns>
    public static string GenerateExcel(EasyExcelBuilder easyExcelBuilderFile, string savePath)
    {
        using var xlWorkbook = ClosedXmlEngine(easyExcelBuilderFile);

        var saveUrl = $"{savePath}\\{easyExcelBuilderFile.FileName}.xlsx";

        // Save
        xlWorkbook.SaveAs(saveUrl);

        return saveUrl;
    }

    /// <summary>
    /// Generate Simple Grid Excel file from special model configured options with EasyExcel attributes
    /// </summary>
    /// <param name="easyGridExcelBuilder"></param>
    /// <returns></returns>
    public static GeneratedExcelFile GenerateGridExcel(EasyGridExcelBuilder easyGridExcelBuilder)
    {
        var easyExcelBuilder = easyGridExcelBuilder.ConvertEasyGridExcelBuilderToEasyExcelBuilder();

        return GenerateExcel(easyExcelBuilder);
    }

    /// <summary>
    /// Generate Simple Grid Excel file from special model configured options with EasyExcel attributes
    /// Save it in path and return the saved url
    /// </summary>
    /// <param name="easyGridExcelBuilder"></param>
    /// <param name="savePath"></param>
    /// <returns></returns>
    public static string GenerateGridExcel(EasyGridExcelBuilder easyGridExcelBuilder, string savePath)
    {
        var easyExcelBuilder = easyGridExcelBuilder.ConvertEasyGridExcelBuilderToEasyExcelBuilder();

        return GenerateExcel(easyExcelBuilder, savePath);
    }

    private static XLWorkbook ClosedXmlEngine(EasyExcelBuilder easyExcelBuilder)
    {
        if (easyExcelBuilder.FileName.IsNullOrWhiteSpace())
            easyExcelBuilder.FileName = $"EasyExcelGeneratedFile_{DateTime.Now:yyyy-MM-dd HH-mm-ss}";

        //-------------------------------------------
        //  Create Workbook (integrated with using statement)
        //-------------------------------------------
        var xlWorkbook = new XLWorkbook
        {
            RightToLeft = easyExcelBuilder.AllSheetsDefaultStyle.AllSheetsDefaultDirection == SheetDirection.RightToLeft,
            ColumnWidth = easyExcelBuilder.AllSheetsDefaultStyle.AllSheetsDefaultColumnWidth,
            RowHeight = easyExcelBuilder.AllSheetsDefaultStyle.AllSheetsDefaultRowHeight
        };

        // Check any sheet available
        if (easyExcelBuilder.Sheets.Count == 0)
            throw new Exception("No sheet is available to create Excel workbook");

        // Check sheet names are unique
        var sheetNames = easyExcelBuilder.Sheets
            .Where(s => s.SheetName.IsNullOrWhiteSpace() is false)
            .Select(s => s.SheetName)
            .ToList();

        var uniqueSheetNames = sheetNames.Distinct().ToList();

        if (sheetNames.Count != uniqueSheetNames.Count)
            throw new Exception("Sheet names should be unique");

        // Auto naming for sheets

        int i = 1;
        foreach (Sheet sheet in easyExcelBuilder.Sheets)
        {
            if (sheet.SheetName.IsNullOrWhiteSpace())
            {
                var isNameOk = false;

                while (isNameOk is false)
                {
                    var possibleName = $"Sheet{i}";

                    isNameOk = easyExcelBuilder.Sheets.Any(s => s.SheetName == possibleName) is false;

                    if (isNameOk)
                        sheet.SheetName = possibleName;

                    i++;
                }
            }
        }

        //-------------------------------------------
        //  Add Sheets one by one to ClosedXML Workbook instance
        //-------------------------------------------
        foreach (var sheet in easyExcelBuilder.Sheets)
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
            var textAlign = sheet.SheetStyle.SheetDefaultTextAlign ?? easyExcelBuilder.AllSheetsDefaultStyle.AllSheetsDefaultTextAlign;

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
                    xlSheet.ConfigureRow(tableRow, sheet.SheetColumnsStyle, sheet.IsSheetLocked ?? easyExcelBuilder.AreSheetsLockedByDefault);
                }

                var tableRange = xlSheet.Range(table.StartCellLocation.Y,
                    table.StartCellLocation.X,
                    table.EndLocation.Y,
                    table.EndLocation.X);

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
                xlSheet.ConfigureRow(sheetRow, sheet.SheetColumnsStyle, sheet.IsSheetLocked ?? easyExcelBuilder.AreSheetsLockedByDefault);
            }

            //-------------------------------------------
            //  Map Cells
            //-------------------------------------------
            foreach (var cell in sheet.SheetCells)
            {
                if (cell.IsCellVisible is false)
                    continue;

                xlSheet.ConfigureCell(cell, sheet.SheetColumnsStyle, sheet.IsSheetLocked ?? easyExcelBuilder.AreSheetsLockedByDefault);
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

    private static EasyExcelBuilder ConvertEasyGridExcelBuilderToEasyExcelBuilder(this EasyGridExcelBuilder easyGridExcelBuilder)
    {
        var easyExcelBuilder = new EasyExcelBuilder();

        foreach (var gridExcelSheet in easyGridExcelBuilder.Sheets)
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

                SheetDirection sheetDirection = SheetDirection.RightToLeft;

                bool isSheetLocked = false;

                ProtectionLevel sheetProtectionLevel = new();

                foreach (var record in records)
                {
                    // Each record is an entire row of Excel

                    var easyExcelSheetAttribute = record.GetType().GetCustomAttribute<ExcelSheetAttribute>();

                    sheetName = easyExcelSheetAttribute?.SheetName;

                    sheetDirection = easyExcelSheetAttribute?.SheetDirection ?? SheetDirection.RightToLeft;

                    var defaultFont = new TextFont
                    {
                        FontName = easyExcelSheetAttribute?.FontName,
                        FontSize = easyExcelSheetAttribute?.FontSize == 0 ? null : easyExcelSheetAttribute?.FontSize,
                        FontColor = Color.FromKnownColor(easyExcelSheetAttribute?.FontColor ?? KnownColor.Black),
                        IsBold = easyExcelSheetAttribute?.IsFontBold
                    };

                    isSheetLocked = easyExcelSheetAttribute?.IsSheetLocked ?? false;

                    var isSheetHardProtected = easyExcelSheetAttribute?.IsSheetHardProtected ?? false;

                    if (isSheetHardProtected)
                        sheetProtectionLevel.HardProtect = true;

                    borderType = easyExcelSheetAttribute?.BorderType ?? LineStyle.Thin;

                    var defaultTextAlign = easyExcelSheetAttribute?.DefaultTextAlign ?? TextAlign.Center;

                    PropertyInfo[] properties = record.GetType().GetProperties();

                    int xLocation = 1;

                    var recordRow = new Row
                    {
                        RowHeight = easyExcelSheetAttribute?.DataRowHeight == 0 ? null : easyExcelSheetAttribute?.DataRowHeight,
                        BackgroundColor = easyExcelSheetAttribute?.DataBackgroundColor != null ? Color.FromKnownColor(easyExcelSheetAttribute.DataBackgroundColor) : Color.Transparent,
                    };

                    foreach (var prop in properties)
                    {
                        var easyExcelColumnAttribute = (ExcelColumnAttribute?)prop.GetCustomAttributes(true).FirstOrDefault(x => x is ExcelColumnAttribute);

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
                            FontName = easyExcelColumnAttribute?.FontName ?? defaultFont.FontName,
                            FontSize = easyExcelColumnAttribute?.FontSize is null || easyExcelColumnAttribute.FontSize == 0 ? defaultFont.FontSize : easyExcelColumnAttribute.FontSize,
                            FontColor = easyExcelColumnAttribute is null || easyExcelColumnAttribute.FontColor == KnownColor.Transparent
                            ? defaultFont.FontColor.Value
                            : Color.FromKnownColor(easyExcelColumnAttribute.FontColor),
                            IsBold = easyExcelColumnAttribute is null || easyExcelColumnAttribute.FontWeight == FontWeight.Inherit
                            ? defaultFont.IsBold
                            : easyExcelColumnAttribute.FontWeight == FontWeight.Bold
                        };

                        // Header
                        if (headerCalculated is false)
                        {
                            headerRow.Cells.Add(new Cell(new CellLocation(xLocation, yLocation))
                            {
                                Value = easyExcelColumnAttribute?.HeaderName ?? prop.Name,
                                CellTextAlign = GetCellTextAlign(defaultTextAlign, easyExcelColumnAttribute?.HeaderTextAlign),
                                CellType = CellType.Text,
                                Font = finalFont
                            });

                            headerRow.RowHeight = easyExcelSheetAttribute?.HeaderHeight == 0 ? null : easyExcelSheetAttribute?.HeaderHeight;

                            headerRow.BackgroundColor = easyExcelSheetAttribute?.HeaderBackgroundColor != null ? Color.FromKnownColor(easyExcelSheetAttribute.HeaderBackgroundColor) : Color.Transparent;

                            headerRow.OutsideBorder = new Border { BorderColor = Color.Black, BorderLineStyle = borderType };

                            headerRow.InsideBorder = new Border { BorderColor = Color.Black, BorderLineStyle = borderType };

                            // Calculate Columns style
                            columnsStyle.Add(new ColumnStyle
                            {
                                ColumnNo = xLocation,
                                ColumnWidth = new ColumnWidth
                                {
                                    Width = easyExcelColumnAttribute?.ColumnWidth == 0 ? null : easyExcelColumnAttribute?.ColumnWidth,
                                    WidthCalculationType = easyExcelColumnAttribute is null || easyExcelColumnAttribute.ColumnWidth == 0 ? ColumnWidthCalculationType.AdjustToContents : ColumnWidthCalculationType.ExplicitValue
                                }
                            });
                        }

                        // Data
                        recordRow.Cells.Add(new Cell(new CellLocation(xLocation, yLocation + 1))
                        {
                            Value = prop.GetValue(record),
                            CellType = easyExcelColumnAttribute?.ExcelDataType ?? CellType.Text,
                            CellTextAlign = GetCellTextAlign(defaultTextAlign, easyExcelColumnAttribute?.DataTextAlign),
                            Font = finalFont
                        });

                        xLocation++;
                    }

                    dataRows.Add(recordRow);

                    yLocation++;

                    headerCalculated = true;
                }

                easyExcelBuilder.Sheets.Add(new Sheet
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

        return easyExcelBuilder;
    }

    private static void ConfigureCell(this IXLWorksheet xlSheet, Cell cell, List<ColumnStyle> columnProps, bool isSheetLocked)
    {
        // Infer XLDataType and value from "cell" CellType
        XLDataType? xlDataType;
        object cellValue = cell.Value;
        switch (cell.CellType)
        {
            case CellType.Number:
                xlDataType = XLDataType.Number;
                break;

            case CellType.Percentage:
                xlDataType = XLDataType.Text;
                cellValue = $"{cellValue}%";
                break;

            case CellType.Currency:
                xlDataType = XLDataType.Number;
                if (cellValue.IsNumber() is false)
                    throw new Exception("Cell with Currency CellType should be Number type");
                cellValue = Convert.ToDecimal(cellValue).ToString("##,###");
                break;

            case CellType.MiladiDate:
                xlDataType = XLDataType.DateTime;
                if (cellValue is not DateTime)
                    throw new Exception("Cell with MiladiDate CellType should be DateTime type");
                break;

            case CellType.Text:
            case CellType.Formula:
                xlDataType = XLDataType.Text;
                break;

            default: // = CellType.General
                xlDataType = null;
                break;
        }

        // Infer XLAlignment from "cell"
        XLAlignmentHorizontalValues? cellAlignmentHorizontalValue = cell.CellTextAlign switch
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
            var x = cell.CellLocation.X;

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
        var locationCell = xlSheet.Cell(cell.CellLocation.Y, cell.CellLocation.X);

        if (xlDataType is not null)
            locationCell.SetDataType((XLDataType)xlDataType);

        if (cell.CellType == CellType.Formula)
            locationCell.SetFormulaA1(cellValue.ToString());
        else
            locationCell.SetValue(cellValue);

        locationCell.Style.Alignment.SetWrapText(cell.Wordwrap);

        locationCell.Style.Protection.Locked = (bool)isLocked;

        if (cellAlignmentHorizontalValue is not null)
            locationCell.Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)cellAlignmentHorizontalValue);

        // Set Vertical Middle Align
        locationCell.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

        // Set Font
        if (cell.Font.FontColor is not null)
            locationCell.Style.Font.SetFontColor(XLColor.FromColor(cell.Font.FontColor.Value));

        if (cell.Font.FontSize is not null)
            locationCell.Style.Font.SetFontSize(cell.Font.FontSize.Value);

        if (cell.Font.IsBold is not null)
            locationCell.Style.Font.SetBold(cell.Font.IsBold.Value);

        if (cell.Font.FontName.IsNullOrWhiteSpace() is false)
            locationCell.Style.Font.SetFontName(cell.Font.FontName);
    }

    private static void ConfigureRow(this IXLWorksheet xlSheet, Row row, List<ColumnStyle> columnsStyleList, bool isSheetLocked)
    {
        foreach (var rowCell in row.Cells)
        {
            if (rowCell.IsCellVisible is false)
                continue;

            xlSheet.ConfigureCell(rowCell, columnsStyleList, isSheetLocked);
        }

        // Configure merged cells in the row
        foreach (var cellsToMerge in row.MergedCellsList)
        {
            // CellsToMerge example is "B2:D2"
            xlSheet.Range(cellsToMerge).Row(1).Merge();
        }

        if (row.Cells.Count != 0)
        {
            if (row.StartCellLocation is not null && row.EndCellLocation is not null)
            {
                var xlRow = xlSheet.Row(row.Cells.First().CellLocation.Y);
                if (row.RowHeight is not null)
                    xlRow.Height = (double)row.RowHeight;

                var xlRowRange = xlSheet.Range(row.StartCellLocation.Y,
                    row.StartCellLocation.X,
                    row.EndCellLocation.Y,
                    row.EndCellLocation.X);

                if (row.Font.FontColor is not null)
                    xlRowRange.Style.Font.SetFontColor(XLColor.FromColor(row.Font.FontColor.Value));

                if (row.Font.FontSize is not null)
                    xlRowRange.Style.Font.SetFontSize(row.Font.FontSize.Value);

                if (row.Font.IsBold is not null)
                    xlRowRange.Style.Font.SetBold(row.Font.IsBold.Value);

                if (row.Font.FontName.IsNullOrWhiteSpace() is false)
                    xlRowRange.Style.Font.SetFontName(row.Font.FontName);

                xlRowRange.Style.Fill.SetBackgroundColor(XLColor.FromColor(row.BackgroundColor));

                XLBorderStyleValues? outsideBorder = GetXlBorderLineStyle(row.OutsideBorder.BorderLineStyle);

                if (outsideBorder is not null)
                {
                    xlRowRange.Style.Border.SetOutsideBorder((XLBorderStyleValues)outsideBorder);
                    xlRowRange.Style.Border.SetOutsideBorderColor(XLColor.FromColor(row.OutsideBorder.BorderColor));
                }

                XLBorderStyleValues? insideBorder = GetXlBorderLineStyle(row.InsideBorder.BorderLineStyle);

                if (insideBorder is not null)
                {
                    xlRowRange.Style.Border.SetInsideBorder((XLBorderStyleValues)insideBorder);
                    xlRowRange.Style.Border.SetInsideBorderColor(XLColor.FromColor(row.InsideBorder.BorderColor));
                }
            }
            else
            {
                var xlRow = xlSheet.Row(row.Cells.First().CellLocation.Y);
                if (row.RowHeight is not null)
                    xlRow.Height = (double)row.RowHeight;

                if (row.Font.FontColor is not null)
                    xlRow.Style.Font.SetFontColor(XLColor.FromColor(row.Font.FontColor.Value));
                if (row.Font.FontSize is not null)
                    xlRow.Style.Font.SetFontSize(row.Font.FontSize.Value);
                if (row.Font.IsBold is not null)
                    xlRow.Style.Font.SetBold(row.Font.IsBold.Value);
                if (row.Font.FontName.IsNullOrWhiteSpace() is false)
                    xlRow.Style.Font.SetFontName(row.Font.FontName);

                xlRow.Style.Fill.SetBackgroundColor(XLColor.FromColor(row.BackgroundColor));
                xlRow.Style.Border.SetOutsideBorder(XLBorderStyleValues.Dotted);
                xlRow.Style.Border.SetInsideBorder(XLBorderStyleValues.Thick);
                xlRow.Style.Border.SetTopBorder(XLBorderStyleValues.Thick);
                xlRow.Style.Border.SetRightBorder(XLBorderStyleValues.DashDotDot);
            }
        }
    }

    private static XLBorderStyleValues? GetXlBorderLineStyle(LineStyle borderLineStyle)
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

    private static bool IsNumber(this object value)
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
}