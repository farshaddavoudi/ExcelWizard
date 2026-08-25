using ClosedXML.Excel;
using ExcelWizard.Models;
using ExcelWizard.Service;

namespace ExcelWizard.IntegrationTests;

internal static class WorkbookTestHelper
{
    public static XLWorkbook GenerateWorkbook(IExcelBuilder builder)
    {
        var service = new ExcelWizardService(new FakeBlazorDownloadFileService());
        var generatedFile = service.GenerateExcel(builder);

        generatedFile.Content.Should().NotBeNullOrEmpty();

        return new XLWorkbook(new MemoryStream(generatedFile.Content!));
    }
}
