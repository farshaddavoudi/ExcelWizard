using ApiApp.SimorghReportModels;
using ExcelWizard.Models;
using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWExcel;
using ExcelWizard.Models.EWMerge;
using ExcelWizard.Models.EWRow;
using ExcelWizard.Models.EWSheet;
using ExcelWizard.Models.EWStyles;
using ExcelWizard.Models.EWTable;
using ExcelWizard.Service;
using System.Drawing;

namespace ApiApp.Service;

public class SimorghExcelBuilderService : ISimorghExcelBuilderService
{
    private readonly IExcelWizardService _excelWizardService;

    // DI
    public SimorghExcelBuilderService(IExcelWizardService excelWizardService)
    {
        _excelWizardService = excelWizardService;
    }


    public GeneratedExcelFile GenerateVoucherStatementExcelReport(VoucherStatementResult voucherStatement)
    {
        IExcelBuilder excelBuilder = GetExcelModelFromVoucherStatementResult(voucherStatement);

        return _excelWizardService.GenerateExcel(excelBuilder);
    }

    public string GenerateVoucherStatementExcelReport(VoucherStatementResult voucherStatement, string savePath)
    {
        var excelBuilder = GetExcelModelFromVoucherStatementResult(voucherStatement);

        return _excelWizardService.GenerateExcel(excelBuilder, savePath);
    }

    /// <summary>
    /// Defined to use for both methods and do not duplicate codes
    /// </summary>
    private IExcelBuilder GetExcelModelFromVoucherStatementResult(VoucherStatementResult voucherStatement)
    {
        // It is the heart of using the ExcelWizard package to generate your desired Excel report
        // You should create your Excel template (ExcelBuilder model) using your local app model (here VoucherStatementResult)
        // Just start with ExcelBuilder and the properties names speak for themselves. Also note all properties
        // have proper comments to make them clear

        // Building Excel Parts: Seeing the Excel template At a glance, we can see it is created of below sections:
        // 1- Top header which is a Table (is not a Row because of occupying two Rows) which is Merged and became a Unite Cell
        // 2- Having a Row which is the first table Header (It can be part of the Table model, but it makes it a little hard because
        // the Table data is dynamic and it is better to see the Table header as a single Row.
        // 3- First table with some dynamic data which the data is currency type
        // 4- Again having a Row which is the second table Header.
        // 5- Second table again with dynamic data.
        // 6- A Table which actually is the bottom data header (with blue bg). It is merged, so cannot be declared as a Row.
        // 7- Bottom table with thick inside border having حقوق پایه و اضافه کار Data in it
        // 8- Last Table for Bottom data which again is merged

        // Table: Excel Header 
        var tableHeader = TableBuilder
            .CreateStepByStepManually()
            .SetRows(RowBuilder
                .SetCells(
                    CellBuilder
                        .SetLocation("A", 1)
                        .SetValue("گزارش تست")
                        .SetCellStyle(new CellStyle
                        {
                            CellTextAlign = TextAlign.Center
                        })
                        .Build()
                    )
                .RowHasNoMerging()
                .RowHasNoCustomStyle()
                .Build())
            .SetTableMergedCells(
                MergeBuilder
                .SetMergingStartPoint("A", 1)
                .SetMergingFinishPoint("H", 2)
                .Build()
                )
            .TableHasNoCustomStyle()
            .Build();

        // Gray bg row (کد حساب - بدهکار - بستانکار) - First table Header
        var rowFirstTableHeader = RowBuilder
            .SetCells(
                CellBuilder.SetLocation("A", 3).SetValue("کد حساب").Build(),
                CellBuilder.SetLocation("B", 3).SetValue("بدهکار").Build(),
                CellBuilder.SetLocation("C", 3).SetValue("بستانکار").Build()
                )
            .RowHasNoMerging()
            .SetRowStyle(new RowStyle
            {
                BackgroundColor = Color.Gray
            })
            .Build();

        // First table with header of (کد حساب - بدهکار - بستانکار)
        var table1St = TableBuilder
            .CreateStepByStepManually()
            .SetRows(voucherStatement.VoucherStatementItem.Select((item, index) =>
                RowBuilder
                    .SetCells(
                        CellBuilder.SetLocation("A", 4).SetValue(item.AccountCode).Build(),
                        CellBuilder.SetLocation("B", 4).SetValue(item.Debit).SetContentType(CellContentType.Currency).Build(),
                        CellBuilder.SetLocation("C", 4).SetValue(item.Credit).SetContentType(CellContentType.Currency).Build()
                        )
                    .RowHasNoMerging()
                    .RowHasNoCustomStyle()
                    .Build()
            ).ToList())
            .TableHasNoMerging()
            .SetTableStyle(new TableStyle
            {
                TableOutsideBorder = new Border { BorderLineStyle = LineStyle.Thick },
                InsideCellsBorder = new Border { BorderLineStyle = LineStyle.Thick }
            })
            .Build();

        // Gray bg row (کد حساب - بدهکار - بستانکار) - Second table Header
        var rowSecondTableHeader = RowBuilder
            .SetCells(
                CellBuilder.SetLocation("A", table1St.GetNextVerticalRowNumberAfterTable()).SetValue("کد حساب").Build(),
                CellBuilder.SetLocation("B", table1St.GetNextVerticalRowNumberAfterTable()).SetValue("بدهکار").Build(),
                CellBuilder.SetLocation("C", table1St.GetNextVerticalRowNumberAfterTable()).SetValue("بستانکار").Build()
            )
            .RowHasNoMerging()
            .SetRowStyle(new RowStyle
            {
                BackgroundColor = Color.Gray
            })
            .Build();

        // Second table with header of (کد حساب - بدهکار - بستانکار)
        var table2Nd = TableBuilder
            .CreateStepByStepManually()
            .SetRows(voucherStatement.VoucherStatementItem.Select((item, index) =>
                RowBuilder
                    .SetCells(
                        CellBuilder
                            .SetLocation("A", index + rowSecondTableHeader.GetNextRowNumberAfterRow())
                            .SetValue(item.AccountCode)
                            .Build(),
                        CellBuilder
                            .SetLocation("B", index + rowSecondTableHeader.GetNextRowNumberAfterRow())
                            .SetValue(item.Debit)
                            .SetContentType(CellContentType.Currency)
                            .Build(),
                        CellBuilder
                            .SetLocation("C", index + rowSecondTableHeader.GetNextRowNumberAfterRow())
                            .SetValue(item.Credit)
                            .SetContentType(CellContentType.Currency)
                            .Build()
                        )
                    .RowHasNoMerging()
                    .RowHasNoCustomStyle()
                    .Build()
            ).ToList())
            .TableHasNoMerging()
            .SetTableStyle(new TableStyle
            {
                TableOutsideBorder = new Border { BorderLineStyle = LineStyle.Thick },
                InsideCellsBorder = new Border { BorderLineStyle = LineStyle.Thick }
            })
            .Build();

        // Bottom section Data Header (Blue Bg)
        // MERGES:
        // The Merges here is very tricky!
        // At first, there are Two Vertical Merges (A17:A18) and (B17:B18). 
        // Then there is one Horizontal Merge (C17:E17) for کارخانه دان
        // The same pattern repeats for پرورش پولت and تخم گزاری جوجه
        // And a Vertical Merge for showing sum (K17:K18) 
        // And a Vertical Merge for Average (L17:L18)
        var tableBottomBlueHeader = TableBuilder
            .CreateStepByStepManually()
            .SetRows(
                RowBuilder
                    .SetCells(
                        CellBuilder.SetLocation("A", table2Nd.GetNextVerticalRowNumberAfterTable()).SetValue("نام حساب")
                            .Build(),
                        CellBuilder.SetLocation("B", table2Nd.GetNextVerticalRowNumberAfterTable()).SetValue("کد حساب")
                            .Build(),
                        CellBuilder.SetLocation("C", table2Nd.GetNextVerticalRowNumberAfterTable())
                            .SetValue("کارخانه دان-51011").Build(),
                        CellBuilder.SetLocation("D", table2Nd.GetNextVerticalRowNumberAfterTable()).Build(),
                        CellBuilder.SetLocation("E", table2Nd.GetNextVerticalRowNumberAfterTable()).Build(),
                        CellBuilder.SetLocation("F", table2Nd.GetNextVerticalRowNumberAfterTable())
                            .SetValue("پرورش پولت-51018").Build(),
                        CellBuilder.SetLocation("G", table2Nd.GetNextVerticalRowNumberAfterTable()).Build(),
                        CellBuilder.SetLocation("H", table2Nd.GetNextVerticalRowNumberAfterTable()).Build(),
                        CellBuilder.SetLocation("I", table2Nd.GetNextVerticalRowNumberAfterTable())
                            .SetValue("تخم گزار تجاری-51035").Build(),
                        CellBuilder.SetLocation("J", table2Nd.GetNextVerticalRowNumberAfterTable()).Build(),
                        CellBuilder.SetLocation("K", table2Nd.GetNextVerticalRowNumberAfterTable()).Build(),
                        CellBuilder.SetLocation("L", table2Nd.GetNextVerticalRowNumberAfterTable())
                            .SetValue("میانگین")
                            .SetCellStyle(new CellStyle
                            {
                                BackgroundColor = Color.White,
                                Font = new TextFont { FontColor = Color.Black }
                            })
                            .Build()
                        )
                    .RowHasNoMerging()
                    .SetRowStyle(new RowStyle { RowHeight = 20 })
                    .Build(),

                RowBuilder
                    .SetCells(
                        CellBuilder.SetLocation("A", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).Build(),
                        CellBuilder.SetLocation("B", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).Build(),
                        CellBuilder.SetLocation("C", table2Nd.GetNextVerticalRowNumberAfterTable() + 1)
                            .SetValue("قبل از تسهیم").Build(),
                        CellBuilder.SetLocation("D", table2Nd.GetNextVerticalRowNumberAfterTable() + 1)
                            .SetValue("بعد از تسهیم").Build(),
                        CellBuilder.SetLocation("E", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).SetValue("جمع")
                            .Build(),
                        CellBuilder.SetLocation("F", table2Nd.GetNextVerticalRowNumberAfterTable() + 1)
                            .SetValue("قبل از تسهیم").Build(),
                        CellBuilder.SetLocation("G", table2Nd.GetNextVerticalRowNumberAfterTable() + 1)
                            .SetValue("بعد از تسهیم").Build(),
                        CellBuilder.SetLocation("H", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).SetValue("جمع")
                            .Build(),
                        CellBuilder.SetLocation("I", table2Nd.GetNextVerticalRowNumberAfterTable() + 1)
                            .SetValue("قبل از تسهیم").Build(),
                        CellBuilder.SetLocation("J", table2Nd.GetNextVerticalRowNumberAfterTable() + 1)
                            .SetValue("بعد از تسهیم").Build(),
                        CellBuilder.SetLocation("K", table2Nd.GetNextVerticalRowNumberAfterTable() + 1).SetValue("جمع")
                            .Build(),
                        CellBuilder
                            .SetLocation("L", table2Nd.GetNextVerticalRowNumberAfterTable() + 1)
                            .SetCellStyle(new CellStyle
                            {
                                BackgroundColor = Color.White
                            }
                            ).Build()
                        )
                    .RowHasNoMerging()
                    .SetRowStyle(new RowStyle { RowHeight = 20 })
                    .Build()
                )
            .SetTableMergedCells(
                MergeBuilder
                    .SetMergingStartPoint("A", table2Nd.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("A", table2Nd.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("B", table2Nd.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("B", table2Nd.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("C", table2Nd.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("E", table2Nd.GetNextVerticalRowNumberAfterTable())
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("F", table2Nd.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("H", table2Nd.GetNextVerticalRowNumberAfterTable())
                    .Build())
            .SetTableStyle(new TableStyle
            {
                BackgroundColor = Color.Blue,
                Font = new TextFont { FontColor = Color.White }
            })
            .Build();

        // Table with Salaries data with thick borders
        var tableSalaries = TableBuilder
            .CreateStepByStepManually()
            .SetRows(voucherStatement.Accounts.Select((account, index) =>
                RowBuilder
                    .SetCells(
                        CellBuilder.SetLocation("A", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + index)
                            .SetValue(account.Name).Build(),
                        CellBuilder.SetLocation("B", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + index)
                            .SetValue(account.Code).Build()
                        )
                    .RowHasNoMerging()
                    .RowHasNoCustomStyle()
                    .Build()
            ).ToList())
            .TableHasNoMerging()
            .SetTableStyle(new TableStyle
            {
                TableOutsideBorder = new Border { BorderLineStyle = LineStyle.Thick, BorderColor = Color.Black },
                InsideCellsBorder = new Border { BorderLineStyle = LineStyle.Thick, BorderColor = Color.Black }
            })
            .Build();

        // Last table with sharing before/after data
        var tableSharingBeforeAfterData = TableBuilder
            .CreateStepByStepManually()
            .SetRows(
                RowBuilder
                    .SetCells(
                        CellBuilder.SetLocation("C", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                            .SetValue(504000).Build(),
                        CellBuilder.SetLocation("D", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                            .SetValue(504000).Build(),
                        CellBuilder.SetLocation("E", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                            .SetValue(504000).Build(),
                        CellBuilder.SetLocation("F", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                            .SetValue(504000).Build(),
                        CellBuilder.SetLocation("G", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                            .SetValue(504000).Build(),
                        CellBuilder.SetLocation("H", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                            .SetValue(504000).Build(),
                        CellBuilder.SetLocation("I", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                            .SetValue(504000).Build(),
                        CellBuilder.SetLocation("J", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                            .SetValue(504000).Build(),
                        CellBuilder.SetLocation("K", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                            .SetValue(504000).Build(),
                        CellBuilder.SetLocation("L", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                            .SetValue(504000).Build()
                        )
                    .RowHasNoMerging()
                    .RowHasNoCustomStyle()
                    .Build()
                )
            .SetTableMergedCells(
                MergeBuilder
                    .SetMergingStartPoint("C", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("C", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("D", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("D", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("E", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("E", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("F", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("F", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("G", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("G", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("H", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("H", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("I", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("I", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("J", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("J", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("K", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("K", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build(),
                MergeBuilder
                    .SetMergingStartPoint("L", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable())
                    .SetMergingFinishPoint("L", tableBottomBlueHeader.GetNextVerticalRowNumberAfterTable() + 1)
                    .Build())
            .SetTableStyle(new TableStyle { TableTextAlign = TextAlign.Center })
            .Build();

        return ExcelBuilder
            .SetGeneratedFileName(voucherStatement.ReportName)
            .CreateComplexLayoutExcel()
            .SetSheets(SheetBuilder
                .SetName("RemainReport")
                .SetTables(
                    tableHeader,
                    table1St,
                    table2Nd,
                    tableBottomBlueHeader,
                    tableSalaries,
                    tableSharingBeforeAfterData
                    )
                .SetRows(rowFirstTableHeader, rowSecondTableHeader)
                .NoMoreTablesRowsOrCells()
                .SheetHasNoCustomStyle()
                .Build())
            .SetSheetsDefaultStyle(new SheetsDefaultStyle
            {
                AllSheetsDefaultTextAlign = TextAlign.Right,
                AllSheetsDefaultDirection = SheetDirection.RightToLeft
            })
            .Build();
    }
}