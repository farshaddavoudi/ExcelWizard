# EasyExcelGenerator

Easily generate Excel file based on a C# model dynamically in a very simple and straightforward way. In addition, make the generated Excel file directly downloadable from Browser without any hassle in case of using Blazor application. The package is a wrapper for ClosedXML and BlazorFileDownload packages.

# How to Use
1- Register EasyExcelGenerator Service in your application (API or Blazor) by using `AddEasyExcelServices` extension. 
```
// Has an optional argument for ServiceLifeTime. The default lifetime is Scoped.
builder.Services.AddEasyExcelServices();
```

2- Inject `IEasyExcelService` into your class and enjoy it!

# How much is it simple to generatne/download Excel with EasyExcelGenerator?

Assuming you have collection of a model e.g. `var users = List<User>()` which you normally get it from database and show them in a Grid or something. Now, 
you decided to have a live Excel report from it anytime you want.

The Model:

```
// The model you want to have Excel report upon it
public class User
{
	public int Id { get; set; }
	public string FullName { get; set; }
	public string PersonnelCode { get; set; }
    public string Nationality { get; set; }
}
```

In your Service or Controller:

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
            new() { Id = 1, FullName = "Ronaldo", PersonnelCode = "980923", Nationality = "Portugal" },
            new() { Id = 2, FullName = "Messi", PersonnelCode = "991126", Nationality = "Argentine" },
            new() { Id = 3, FullName = "Mbappe", PersonnelCode = "991213", Nationality = "France" }
        };

        // Below will create Excel file as byte[] data
        // Just passing your data to method argument and let the rest to the package! hoorya!
        // This method has an optional parameter `generatedFileName` which is obvious by the name
        byte[] excelFileAsByteArray = _easyExcelService.GenerateGridLayoutExcel(myUsers);

        // Below will create Excel file in specified path and return the full path as string
        // The last param is generated file name
        string fullPathAsString = _easyExcelService.GenerateGridLayoutExcel(myUsers, @"C:\GeneratedExcelSamples", "Users-Excel");

        return Ok(result);
    }
}

```

In case you are coding in Blazor application, the scenario is even simpler. Only get the raw data (=`myUsers`) from API and use the `BlazorDownloadGridLayoutExcel` method
of `EasyExcelService`, the Excel file will be instantly downloaded (by opening download popup) from the browser without any struggle for byte[] handling or something :)

In IndexPage.razor:

```
<button @onclick="DownloadExcelReport"> Export Excel </button>

@code {

    // Inject the Service 
    [Inject] private IEasyExcelService EasyExcelService { get; set; } = default!;

    private async Task<DownloadFileResult> DownloadExcelReport()
    {
        // Get your data from API usually by Http call
        var myUsers = await apiService.GetMyUsers();

        // Just pass the data to method and you are good to go ;)
        // This method has an optional parameter `generatedFileName` which is obvious by the name
        return EasyExcelService.BlazorDownloadGridLayoutExcel(myUsers);
    }
}
```
