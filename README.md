# EasyExcelGenerator

Easily generate Excel file based on a C# model dynamically in a very simple and straightforward way. In addition, make the generated Excel file directly downloadable from Browser without any hassle in case of using Blazor application. The package is a wrapper for ClosedXML and BlazorFileDownload packages.

# How much is it simple to generatne/export Excel with EasyExcelGenerator?

Assuming you have collection of a model e.g. `var users = List<User>()` which you normally get it from database and show them in a Grid or something. Now, 
you decided to have a live Excel report from it anytime you want.

```
// The model you want to have Excel report upon it
public class User
{
	public int Id { get; set; }
	public string FullName { get; set; }
	public string PersonnelCode { get; set; }
}
```

```
public class UserController : ControllerBase 
{
    // Inject the IEasyExcelService Service 
	private IEasyExcelService _easyExcelService;

	public ExcelController(IEasyExcelService easyExcelService)
    {
        _easyExcelService = easyExcelService;
    }

	[HttpGet("export-users")]
    public IActionResult ExportUsersToExcel()
    {
        // The below data normally comes from your database
        // Show static for demo purposes
        var myUsers = new List<User>
        { 
            new() { Id = 1, FullName = "Ronaldo", PersonnelCode = "980923" },
            new() { Id = 2, FullName = "Messi", PersonnelCode = "991126" },
            new() { Id = 2, FullName = "Mbappe", PersonnelCode = "991213" }
        };

        // Below will create Excel file as byte[] data
        byte[] excelFileAsByteArray = _easyExcelService.GenerateGridLayoutExcel(myUsers);

        // Below will create Excel file in specified path and return the total path as string
        string savedPathAsString = _easyExcelService.GenerateGridLayoutExcel(myUsers, @"C:\GeneratedExcelSamples

        return Ok(result);
    }
}

```
