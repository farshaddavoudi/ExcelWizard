# ExcelWizard
[![NuGet Version](https://img.shields.io/nuget/v/ExcelWizard.svg?style=flat)](https://www.nuget.org/packages/ExcelWizard/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://raw.githubusercontent.com/farshaddavoudi/ExcelWizard/master/LICENSE)

Using ExcelWizard, you can easily generate Excel file in a very simple and straightforward way. In addition, make the generated Excel file directly downloadable from Browser without any hassle in case of using Blazor application. The package is a wrapper for ClosedXML and BlazorFileDownload packages.

# How to Use
1. Install `ExcelWizard` pacakge from nuget package manager.

2. Register ExcelWizard Service in your application (API or Blazor) by using the AddExcelWizardServices extension.
```csharp
// Has a `isBlazorApp` argument (default is `false`). 
// In case of using in Blazor application, pass the true value to register necessary services.
// Has an optional argument for ServiceLifeTime. The default lifetime is Scoped.
builder.Services.AddExcelWizardServices(isBlazorApp: false);
```

3. Inject `IExcelWizardService` into your class and enjoy it!

# How much is it simple to generate/download Excel with ExcelWizard?

Assuming you have a collection of a model e.g. `var users = List<User>()` which you normally get it from a database and show them in a Grid or something. Now, 
you decided to have a live Excel report from it anytime you want. Something like below:

<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/Screenshot-5.png">

The Model:

```csharp
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

```csharp
public class UserController : ControllerBase 
{
    // Inject the IExcelWizardService Service 
    private IExcelWizardService _excelWizardService;

    public ExcelController(IExcelWizardService excelWizardService)
    {
        _excelWizardService = excelWizardService;
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
        GeneratedExcelFile generatedExcelFile = _excelWizardService.GenerateGridLayoutExcel(myUsers);

        // Below will create Excel file in specified path and return the full path as string
        // The last param is generated file name
        string fullPathAsString = _excelWizardService.GenerateGridLayoutExcel(myUsers, @"C:\GeneratedExcelSamples", "Users-Excel");

        return Ok(generatedExcelFile);
    }
}

```

In case you are coding in the Blazor application, the scenario is even simpler. Only get the raw data (=`myUsers`) from API and use the `BlazorDownloadGridLayoutExcel` method
of `ExcelWizardService`, the Excel file will be instantly downloaded (by opening the download popup) from the browser without any struggle for byte[] handling or something :)

In IndexPage.razor:

```razor
<button @onclick="DownloadExcelReport"> Export Excel </button>

@code {

    // Inject the Service 
    [Inject] private IExcelWizardService ExcelWizardService { get; set; } = default!;

    private async Task<DownloadFileResult> DownloadExcelReport()
    {
        // Get your data from API usually by Http call
        var myUsers = await apiService.GetMyUsers();

        // Just pass the data to method and you are good to go ;)
        // This method has an optional parameter `generatedFileName` which is obvious by the name
        return ExcelWizardService.BlazorDownloadGridLayoutExcel(myUsers);
    }
}
```

# Concepts

### The Excel you want can be two types: 

1. **Grid-layout** like data; meaning you have a list of data (again like `myUsers`) and you want to easily export it to Excel. The Excel would be 
relatively simple, having a table-like layout, a header, and data. The first examples in the doc were from this type.

<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/Screenshot-3.png">

2. **Compound Excel**; a little more complex than the previous Grid-layout one. This Excel type can include some different Rows, Tables, and special Cells each placed
in different Excel locations. The first type is easier and most straightforward and this type has a different Excel build scenario (Using `GenerateCompoundExcel` method of `IExcelWizardService`).

<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/Screenshot-4.png">

### Also, you can have different scenario in saving or retrieving generated Excel file:

1. Get the *`byte[]`* of the Excel file and use it for your use case, e.g. sending to another client to be shown or saving in a database, etc.

2. Save the Excel directly *on disk* and get the full path address to send to the app client or save it in the database.

3. (Blazor app) Normally you want to *show the Excel to the user as exported file and do not want to save it somewhere*. If your app client is 
something other than Blazor (e.g. React, Angular, or MVC), your only choice is to work with generated Excel byte[] data and handle it for 
the result you want, but for Blazor apps the story is very simple. Just use the `BlazorDownloadGridLayoutExcel` and `BlazorDownloadCompoundExcel` methods *(notice the methods name start with BlazorDownload)* from `IExcelWizardService` in some click event and the Excel file will be generated and instantly downloaded *(by opening the download popup)* right from 
the browser. Easy-peasy, huh! :)

Knowing these simple concepts,  you will easily understand the IExcelWizardService methods and will be able to generate your favorable Excel
in a very easy and fast way.

**`IExcelWizardService` all Methods at a glance:**
```csharp
public interface IExcelWizardService
{
    // Generate Simple Grid Excel file from special model configured options with ExcelWizard attributes
    public GeneratedExcelFile GenerateGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder);

    // Generate Simple Single Sheet Grid Excel file from special model configured options with ExcelWizard attributes
    public GeneratedExcelFile GenerateGridLayoutExcel(object singleSheetDataList, string? generatedFileName = null);

    // Generate Grid-Layout Excel having multiple Sheets from special model configured options with ExcelWizard attributes
    public string GenerateGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder, string savePath);

    // Generate Simple Single Sheet Grid Excel file from special model configured options with ExcelWizard attributes
    public string GenerateGridLayoutExcel(object singleSheetDataList, string savePath, string generatedFileName);

    // Generate Compound Excel consisting multiple parts like some Rows, Tables, Special Cells, etc each in different Excel Location
    public GeneratedExcelFile GenerateCompoundExcel(CompoundExcelBuilder compoundExcelBuilder);

    // Generate Excel file and save it in path and return the saved url
    public string GenerateCompoundExcel(CompoundExcelBuilder compoundExcelBuilder, string savePath);

    #region Blazor Application

    // [Blazor only] Generate and Download instantly from Browser the Simple Multiple Sheet Grid Excel file from special model configured options with ExcelWizard attributes in Blazor apps
    public Task<DownloadFileResult> BlazorDownloadGridLayoutExcel(GridLayoutExcelBuilder multiSheetsGridLayoutExcelBuilder);

    // [Blazor only] Generate and Download instantly from Browser the Simple Single Sheet Grid Excel file from special model configured options with ExcelWizard attributes in Blazor apps
    public Task<DownloadFileResult> BlazorDownloadGridLayoutExcel(object singleSheetDataList, string? generatedFileName = null);

    // [Blazor only] Generate and Download instantly from Browser the Compound Excel consisting multiple parts like some Rows, Tables, Special Cells, etc each in different Excel Location in Blazor apps
    public Task<DownloadFileResult> BlazorDownloadCompoundExcel(CompoundExcelBuilder compoundExcelBuilder);

    #endregion
}
```

# Generate *Single* Sheet Grid-Layout Excel In Details

For single sheet Grid-layout Excel, it is as easy as passing the data (collection of a model like `var myUsers = new List<User>();`, remember?)
to the `GenerateGridLayoutExcel` method (or `BlazorDownloadGridLayoutExcel` in the case of the Blazor app). It will generate (download in Blazor) a very
simple Excel filled with data without any Excel customization. 

<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/Screenshot-5.png" />

You saw a simple example in <i>how much is it simple section</i>. Below we are just adding a variable (`IsLoading`) to show some loading during fetching data and generating Excel:

```razor
<button @onclick="DownloadExcelReport"> Export Excel </button>

@code {

    // Inject the Service 
    [Inject] private IExcelWizardService ExcelWizardService { get; set; } = default!;

    private async Task<DownloadFileResult> DownloadExcelReport()
    {
        // IsLoading is a hypothetical variable to show some Loading in your page
        IsLoading = true;
        
        try {
            // Get your data from API usually by Http call
            var myUsers = await apiService.GetMyUsers();

            // Just pass the data to method and you are good to go ;)
            // This method has an optional parameter `generatedFileName` which is obvious by the name
            return await ExcelWizardService.BlazorDownloadGridLayoutExcel(myUsers);
        }
        finally {
            // Finish Showing loading 
            IsLoading = false;
        }
    }
}
```

### Customize Excel using **`[ExcelSheet]`** and **`[ExcelColumn]`** Attributes
For example, ignore a column (property) to be shown in exported Excel, having some aligns for header or cells, text font/size/color, different background color for header or cells or a specific column!,
custom header name for a column, custom header height or column width or Sheet direction (RTL/LTR), etc..? All these options plus a lot more can be configured by two
attributes you can use on your model. `[ExcelSheet]` for Excel generic properties and `[ExcelColumn]` for per property (column) customization.

Remember the User model, we can use the attributes below to customize our Users Excel:

```csharp
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
    
    [ExcelColumn(Ignore = true)]
    public string? Age { get; set; }
}
```

The Result:
<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/Screenshot-2.png" />

You do not need to remember all the properties. Just use the attribute and intellisense will show you all the available options you 
can use to customize Excel.

# Generate *Multiple* Sheets Grid-Layout Excel In Details

It is almost the same with a single sheet, using the second overload of the same `GenerateGridLayoutExcel` method (`BlazorDownloadGridLayoutExcel` method in case of the Blazor app) with the `GridLayoutExcelBuilder` argument 
 that should be provided to configure the Excel file. The Excel customization 
is exactly like Single sheet Grid-layout Excel (See the previous section).

# Generate *Compound* Excel
Generating Excel in this case for single or multi sheets are the same. Using the `GenerateCompoundExcel` method (`BlazorDownloadCompoundExcel` in case of Blazor application) you can create any customized Excel file. Just go along with the `CompoundExcelBuilder` argument and provide the necessary parts for your Excel.

Tip: We do not use any attributes (`[ExcelSheet]` and `[ExcelColumn]`) here.

### Complete Example of Building Compound Excel from Scratch
Let's assume we have an application related to a company financials and we want to have a custom Excel report like below format:

<img src="https://github.com/farshaddavoudi/ExcelWizard/blob/main/screenshots/accounts-excel-template-report.png" />

You can download the Excel from here.

So, clearly it is a Compound Excel, not the GridLayout one. In this case, we should fetch all the data from our database, in other words, get the report database model (DTO) to show them to our Excel.

### Steps to create `CompoundExcelBuilder` model:

**
#####1- Analyze Excel Template and Divide It to Sections
#####2- Create Sections Model
#####3- Create the `CompoundExcelBuilder` model
#####4- Use `ExcelWizardService` and `CompoundExcelBuilder` model to generate our Excel
**

**1- Analyze Excel Template and Divide It to Sections**

Analyze the Excel template and divide it to **Table**s, **Row**s and **Cell**s sections (Each section will be mapped to its ExcelWizard model), then create each section model. We use these section models in the next step to create the `CompoundExcelBuilder` model by putting them all together.

For our example, the Excel is composed of these sections:

1- 
2- 
3- 
4- 
5- 

**2- Create each section model**

These models are `Table` model, `Row` model and `Cell` model. All of these models are ExcelWizard models and will use in generating the main `CompoundExcelBuilder` model.

**3- Create the `ExcelBuilderBuilder`**

At last, we create our main model by using the sections model created in Step 2 plus other styles that are available in this class.