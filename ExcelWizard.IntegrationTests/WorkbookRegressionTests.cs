using System.Drawing;
using ClosedXML.Excel;
using ExcelWizard.Models;
using ExcelWizard.Models.EWCell;
using ExcelWizard.Models.EWExcel;
using ExcelWizard.Models.EWMerge;
using ExcelWizard.Models.EWRow;
using ExcelWizard.Models.EWSheet;
using ExcelWizard.Models.EWStyles;
using ExcelWizard.Service;

namespace ExcelWizard.IntegrationTests;

public class WorkbookRegressionTests
{
    [Fact]
    public void GeneratedFile_PreservesRequestedMetadataAndCanBeReopened()
    {
        var service = new ExcelWizardService(new FakeBlazorDownloadFileService());
        var file = service.GenerateExcel(BuildComplexWorkbook());

        file.FileName.Should().Be("regression-workbook");
        file.Extension.Should().Be("xlsx");
        file.MimeType.Should().Be("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        file.Content.Should().NotBeNullOrEmpty();

        using var workbook = new XLWorkbook(new MemoryStream(file.Content!));
        workbook.Worksheets.Count.Should().Be(2);
    }

    [Fact]
    public void ComplexWorkbook_RoundTripsMergesStylesDirectionAndProtection()
    {
        using var workbook = WorkbookTestHelper.GenerateWorkbook(BuildComplexWorkbook());
        var worksheet = workbook.Worksheet("Protected RTL");

        worksheet.RightToLeft.Should().BeTrue();
        worksheet.Protection.IsProtected.Should().BeTrue();
        worksheet.Range("A1:B1").IsMerged().Should().BeTrue();
        worksheet.Cell("A1").GetString().Should().Be("Merged title");
        worksheet.Cell("A1").Style.Fill.BackgroundColor.Color.ToArgb().Should().Be(Color.LightBlue.ToArgb());
        worksheet.Cell("A2").Style.Border.LeftBorder.Should().Be(XLBorderStyleValues.Thick);
        workbook.Worksheet("Second Sheet").Cell("A1").GetString().Should().Be("second");
    }

    private static IExcelBuilder BuildComplexWorkbook()
    {
        var titleRow = RowBuilder
            .SetCells(
                CellBuilder.SetLocation("A", 1).SetValue("Merged title").Build(),
                CellBuilder.SetLocation("B", 1).Build())
            .RowHasNoMerging()
            .SetRowStyle(new RowStyle { BackgroundColor = Color.LightBlue })
            .Build();
        var dataRow = RowBuilder
            .SetCells(
                CellBuilder.SetLocation("A", 2)
                    .SetValue("bordered")
                    .SetCellStyle(new CellStyle
                    {
                        CellBorder = new Border(LineStyle.Thick, Color.DarkBlue)
                    })
                    .Build())
            .RowHasNoMerging()
            .RowHasNoCustomStyle()
            .Build();
        var merge = MergeBuilder
            .SetMergingStartPoint("A", 1)
            .SetMergingFinishPoint("B", 1)
            .SetMergingAreaBackgroundColor(Color.LightBlue)
            .Build();
        var protectedSheet = SheetBuilder
            .SetName("Protected RTL")
            .SetRows(titleRow, dataRow)
            .NoMoreTablesRowsOrCells()
            .SetSheetStyle(new SheetStyle { SheetDirection = SheetDirection.RightToLeft })
            .SetMergedCells(merge)
            .SetSheetLocked(true)
            .SetSheetProtected()
            .SetProtectionLevel(new ProtectionLevel { Password = "test", HardProtect = true })
            .Build();
        var secondSheet = SheetBuilder
            .SetName("Second Sheet")
            .SetCells(CellBuilder.SetLocation("A", 1).SetValue("second").Build())
            .NoMoreTablesRowsOrCells()
            .SheetHasNoCustomStyle()
            .Build();

        return ExcelBuilder
            .SetGeneratedFileName("regression-workbook")
            .CreateComplexLayoutExcel()
            .SetSheets(protectedSheet, secondSheet)
            .SheetsHaveNoDefaultStyle()
            .SetDefaultLockedStatus(false)
            .Build();
    }
}
