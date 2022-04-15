using ClosedXML.Excel;
using ClosedXML.Report.Utils;
using EasyExcelGenerator.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EasyExcelGenerator.Service;

// TODO: Remove static and make them work with DI
public static class EasyExcelService
{
    private static XLWorkbook ClosedXmlEngine(ExcelModel excelModel)
    {
        //-------------------------------------------
        //  Create Workbook (integrated with using statement)
        //-------------------------------------------
        var xlWorkbook = new XLWorkbook
        {
            RightToLeft = excelModel.AllSheetsDefaultStyle.AllSheetsDefaultDirection == SheetDirection.RightToLeft,
            ColumnWidth = excelModel.AllSheetsDefaultStyle.AllSheetsDefaultColumnWidth,
            RowHeight = excelModel.AllSheetsDefaultStyle.AllSheetsDefaultRowHeight
        };

        // Check any sheet available
        if (excelModel.Sheets.Count == 0)
            throw new Exception("No sheet is available to create Excel workbook");

        // Check sheet names are unique
        var sheetNames = excelModel.Sheets
            .Where(s => s.SheetName.IsNullOrWhiteSpace() is false)
            .Select(s => s.SheetName)
            .ToList();

        var uniqueSheetNames = sheetNames.Distinct().ToList();

        if (sheetNames.Count != uniqueSheetNames.Count)
            throw new Exception("Sheet names should be unique");

        // Auto naming for sheets

        int i = 1;
        foreach (Sheet sheet in excelModel.Sheets)
        {
            if (sheet.SheetName.IsNullOrWhiteSpace())
            {
                var isNameOk = false;

                while (isNameOk is false)
                {
                    var possibleName = $"Sheet{i}";

                    isNameOk = excelModel.Sheets.Any(s => s.SheetName == possibleName) is false;

                    if (isNameOk)
                        sheet.SheetName = possibleName;

                    i++;
                }
            }
        }

        //-------------------------------------------
        //  Add Sheets one by one to ClosedXML Workbook instance
        //-------------------------------------------
        foreach (var sheet in excelModel.Sheets)
        {
            // Set name
            var xlSheet = xlWorkbook.Worksheets.Add(sheet.SheetName);

            // Set protection level
            var protection = xlSheet.Protect(sheet.SheetProtectionLevels.Password);
            if (sheet.SheetProtectionLevels.DeleteColumns)
                protection.Protect().AllowedElements = XLSheetProtectionElements.DeleteColumns;
            if (sheet.SheetProtectionLevels.EditObjects)
                protection.Protect().AllowedElements = XLSheetProtectionElements.EditObjects;
            if (sheet.SheetProtectionLevels.FormatCells)
                protection.Protect().AllowedElements = XLSheetProtectionElements.FormatCells;
            if (sheet.SheetProtectionLevels.FormatColumns)
                protection.Protect().AllowedElements = XLSheetProtectionElements.FormatColumns;
            if (sheet.SheetProtectionLevels.FormatRows)
                protection.Protect().AllowedElements = XLSheetProtectionElements.FormatRows;
            if (sheet.SheetProtectionLevels.InsertColumns)
                protection.Protect().AllowedElements = XLSheetProtectionElements.InsertColumns;
            if (sheet.SheetProtectionLevels.InsertHyperLinks)
                protection.Protect().AllowedElements = XLSheetProtectionElements.InsertHyperlinks;
            if (sheet.SheetProtectionLevels.InsertRows)
                protection.Protect().AllowedElements = XLSheetProtectionElements.InsertRows;
            if (sheet.SheetProtectionLevels.SelectLockedCells)
                protection.Protect().AllowedElements = XLSheetProtectionElements.SelectLockedCells;
            if (sheet.SheetProtectionLevels.DeleteRows)
                protection.Protect().AllowedElements = XLSheetProtectionElements.DeleteRows;
            if (sheet.SheetProtectionLevels.EditScenarios)
                protection.Protect().AllowedElements = XLSheetProtectionElements.EditScenarios;
            if (sheet.SheetProtectionLevels.SelectUnlockedCells)
                protection.Protect().AllowedElements = XLSheetProtectionElements.SelectUnlockedCells;
            if (sheet.SheetProtectionLevels.Sort)
                protection.Protect().AllowedElements = XLSheetProtectionElements.Sort;
            if (sheet.SheetProtectionLevels.UseAutoFilter)
                protection.Protect().AllowedElements = XLSheetProtectionElements.AutoFilter;
            if (sheet.SheetProtectionLevels.UsePivotTableReports)
                protection.Protect().AllowedElements = XLSheetProtectionElements.PivotTables;

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
            var textAlign = sheet.SheetStyle.SheetDefaultTextAlign ?? excelModel.AllSheetsDefaultStyle.AllSheetsDefaultTextAlign;

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
            foreach (var columnStyle in sheet.SheetColumnsStyleList)
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
                    xlSheet.ConfigureRow(tableRow, sheet.SheetColumnsStyleList, sheet.IsSheetLocked ?? excelModel.AreSheetsLockedByDefault);
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
                xlSheet.ConfigureRow(sheetRow, sheet.SheetColumnsStyleList, sheet.IsSheetLocked ?? excelModel.AreSheetsLockedByDefault);
            }

            //-------------------------------------------
            //  Map Cells
            //-------------------------------------------
            foreach (var cell in sheet.SheetCells)
            {
                if (cell.Visible is false)
                    continue;

                xlSheet.ConfigureCell(cell, sheet.SheetColumnsStyleList, sheet.IsSheetLocked ?? excelModel.AreSheetsLockedByDefault);
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
        XLAlignmentHorizontalValues? cellAlignmentHorizontalValue = cell.TextAlign switch
        {
            TextAlign.Center => XLAlignmentHorizontalValues.Center,
            TextAlign.Left => XLAlignmentHorizontalValues.Left,
            TextAlign.Right => XLAlignmentHorizontalValues.Right,
            TextAlign.Justify => XLAlignmentHorizontalValues.Justify,
            _ => null
        };

        // Get IsLocked property based on Sheet and Cell "IsLocked" prop
        bool? isLocked = cell.IsLocked;

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

        locationCell.Style
            .Alignment.SetWrapText(cell.Wordwrap);

        locationCell.Style.Protection.SetLocked((bool)isLocked!);

        if (cellAlignmentHorizontalValue is not null)
            locationCell.Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)cellAlignmentHorizontalValue!);
    }

    private static void ConfigureRow(this IXLWorksheet xlSheet, Row row, List<ColumnStyle> columnsStyleList, bool isSheetLocked)
    {
        foreach (var rowCell in row.Cells)
        {
            if (rowCell.Visible is false)
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

                var xlRowRange = xlSheet.Range(row.StartCellLocation.Y, row.StartCellLocation.X, row.EndCellLocation.Y,
                    row.EndCellLocation.X);
                xlRowRange.Style.Font.SetFontColor(XLColor.FromColor(row.FontColor));
                xlRowRange.Style.Fill.SetBackgroundColor(XLColor.FromColor(row.BackgroundColor));

                XLBorderStyleValues? outsideBorder = GetXlBorderLineStyle(row.OutsideBorder.BorderLineStyle);

                if (outsideBorder is not null)
                {
                    xlRowRange.Style.Border.SetOutsideBorder((XLBorderStyleValues)outsideBorder);
                    xlRowRange.Style.Border.SetOutsideBorderColor(
                        XLColor.FromColor(row.OutsideBorder.BorderColor));
                }

                // TODO: For Inside border, the row should be considered as Ranged (like Table). I persume it is not important for this phase
            }
            else
            {
                var xlRow = xlSheet.Row(row.Cells.First().CellLocation.Y);
                if (row.RowHeight is not null)
                    xlRow.Height = (double)row.RowHeight;
                xlRow.Style.Font.SetFontColor(XLColor.FromColor(row.FontColor));
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