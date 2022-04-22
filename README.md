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
	public string? FullName { get; set; }
	public string? PersonnelCode { get; set; }
    public string? Nationality { get; set; }
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
# Concepts Short Explanations

The Excel you want can be two types: 

1- Grid layout like data; meaning you have list of data (again like `myUsers`) and you want easily export it to Excel. The Excel would be 
relativly simple, having table like layout, a header and data. The first examples in the doc was from this type.

2- Compound Excel; meaning more complex than the previous Grid layout one. This Excel can includes some different Rows, Tables, special Cells each placed
in different Excel locations. Clearly, the first type is easier and most straight-forward and this type have different Excel build scenario (Using `GenerateCompoundExcel` method of `IEasyExcelService`).

Also your can have different scenario in saving/reteriving generated Excel file:

1- Get the byte[] of Excel file and use it for your use case, e.g. sending to another client to be shown or saving in database, etc.

2- Save the Excel directly in disk and get the full path address to send to client or save in database.

3- (Blazor app) Normally you want to show the Excel to the user as exported file and do not want to save it somewhere. If your app client is 
something other than Blazor (e.g. React, Angular or MVC), your only choice is working with generated Excel byte[] data and handle it for 
the result you want, but for Blazor apps the story is very simple. Just use the `BlazorDownloadGridLayoutExcel` and `BlazorDownloadCompoundExcel` methods
from `IEasyExcelService` in some click event and the Excel file will be generted and instantly downloaded (by opening download popup) right from 
the browser. Easy-peasy, huh! :)

Knowing these simple concepts,  you will easily understand the IEasyExcelService methods and will be enable to generate your favorable Excel
in a very easy and fast way.

IEasyExcelService Methods:
```
public interface IEasyExcelService
{
    /// <summary>
    /// Generate Simple Grid Excel file from special model configured options with EasyExcel attributes
    /// </summary>
    /// <param name="multiSheetsGridLayoutExcelBuilder"></param>
    /// <returns></returns>
    public GeneratedExcelFile GenerateGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder);

    /// <summary>
    /// Generate Simple Single Sheet Grid Excel file from special model configured options with EasyExcel attributes
    /// </summary>
    /// <param name="singleSheetDataList"> List of data (should be something like List{Person}()) </param>
    /// <param name="generatedFileName"> Generated file name. If leave empty, automatically will have a name like EasyExcelGeneratedFile_2022-04-22 14-06-29 </param>
    /// <returns></returns>
    public GeneratedExcelFile GenerateGridLayoutExcel(object singleSheetDataList, string? generatedFileName = null);

    /// <summary>
    /// Generate Grid Layout Excel having multiple Sheets from special model configured options with EasyExcel attributes
    /// Save it in path and return the saved url
    /// </summary>
    /// <param name="multiSheetsGridLayoutExcelBuilder"> Model for Multiple Sheets Grid Layout Excel. For Single Sheet, use another overload with object arg </param>
    /// <param name="savePath"></param>
    /// <returns></returns>
    public string GenerateGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder, string savePath);

    /// <summary>
    /// Generate Simple Single Sheet Grid Excel file from special model configured options with EasyExcel attributes
    /// Save it in path and return the saved url
    /// </summary>
    /// <param name="singleSheetDataList"> List of data (should be something like List{Person}()) </param>
    /// <param name="savePath"></param>
    /// <param name="generatedFileName"> Generated file name. If leave empty string, automatically will have a name like EasyExcelGeneratedFile_2022-04-22 14-06-29 </param>
    /// <returns></returns>
    public string GenerateGridLayoutExcel(object singleSheetDataList, string savePath, string generatedFileName);

    /// <summary>
    /// Generate Compound Excel consisting multiple parts like some Rows, Tables, Special Cells, etc each in different Excel Location
    /// </summary>
    /// <param name="compoundExcelBuilder"></param>
    /// <returns></returns>
    public GeneratedExcelFile GenerateCompoundExcel(CompoundExcelBuilder compoundExcelBuilder);

    /// <summary>
    /// Generate Excel file and save it in path and return the saved url
    /// </summary>
    /// <param name="compoundExcelBuilder"></param>
    /// <param name="savePath"></param>
    /// <returns></returns>
    public string GenerateCompoundExcel(CompoundExcelBuilder compoundExcelBuilder, string savePath);


    #region Blazor Application

    /// <summary>
    /// [Blazor only] Generate and Download instantly from Browser the Simple Multiple Sheet Grid Excel file from special model configured options with EasyExcel attributes in Blazor apps
    /// </summary>
    /// <param name="multiSheetsGridLayoutExcelBuilder"></param>
    /// <returns></returns>
    public Task<DownloadFileResult> BlazorDownloadGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder);

    /// <summary>
    /// [Blazor only] Generate and Download instantly from Browser the Simple Single Sheet Grid Excel file from special model configured options with EasyExcel attributes in Blazor apps
    /// </summary>
    /// <param name="singleSheetDataList"> List of data (should be something like List{Person}()) </param>
    /// <param name="generatedFileName"> Generated file name. If leave empty, automatically will have a name like EasyExcelGeneratedFile_2022-04-22 14-06-29 </param>
    /// <returns></returns>
    public Task<DownloadFileResult> BlazorDownloadGridLayoutExcel(object singleSheetDataList, string? generatedFileName = null);

    /// <summary>
    /// [Blazor only] Generate and Download instantly from Browser the Compound Excel consisting multiple parts like some Rows, Tables, Special Cells, etc each in different Excel Location in Blazor apps
    /// </summary>
    /// <param name="compoundExcelBuilder"></param>
    /// <returns></returns>
    public Task<DownloadFileResult> BlazorDownloadCompoundExcel(CompoundExcelBuilder compoundExcelBuilder);

    #endregion
}
```

# Generate Single Sheet Grid Layout Excel

For single sheet Grid layout Excel, it is as easy as passing the data (collection of a model like `var myUsers = new List<User>();`, remember?)
to the `GenerateGridLayoutExcel` method (or `BlazorDownloadGridLayoutExcel` in case of Blazor app). It will generate (download in Blazor) a very
simple Excel filled with data without any Excel customization. 

You see the example in how much is it simple section.

## What if you want some customization for the generated Excel file?
For example having some aligns for header or cells, text font/size/color, different background color for header or cells or a specific column!,
custom header name for a column, custom header height or column width and etc..? All these options plus a lot more can be configured by two
attributes you can use on your model. `[ExcelSheet]` for Excel generic properties and `[ExcelColumn]` for per property (column) customization.

Remember the User model, we can use the attributes like below to customize our Users Excel:

```
[ExcelSheet(SheetName = "MyUsers", DefaultTextAlign = TextAlign.Center, HeaderBackgroundColor = KnownColor.LightBlue, HeaderHeight = 40,
    BorderType = LineStyle.DashDotDot, DataBackgroundColor = KnownColor.Bisque, DataRowHeight = 25, IsSheetLocked = true,
    SheetDirection = SheetDirection.RightToLeft, FontColor = KnownColor.Red)]
public class User
{
    [ExcelColumn(HeaderName = "UserId", HeaderTextAlign = TextAlign.Right, DataTextAlign = TextAlign.Right, FontColor = KnownColor.Blue)]
    public int Id { get; set; }

    [ExcelColumn(HeaderName = "Name", HeaderTextAlign = TextAlign.Left, FontWeight = FontWeight.Bold)]
    public string? FullName { get; set; }

    [ExcelColumn(HeaderName = "Personnel No", HeaderTextAlign = TextAlign.Left, ColumnWidth = 50, FontSize = 15)]
    public string? PersonnelCode { get; set; }

    public string? Nationality { get; set; }
}
```

You do not need to remember all the properties. Just use the attribute and intellisense will show you all the available fields you 
can use to customize the Excel.