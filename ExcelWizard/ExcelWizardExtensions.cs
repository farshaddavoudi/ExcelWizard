using BlazorDownloadFile;
using ExcelWizard.Service;
using Microsoft.Extensions.DependencyInjection;
using System;

namespace ExcelWizard;

public static class ExcelWizardExtensions
{
    /// <summary>
    /// Add ExcelWizard package required services to IServiceCollection
    /// </summary>
    /// <param name="services"></param>
    /// <param name="isBlazorApp"> Do you register for a Blazor app or not. In case of API or MVC project, is will be false </param>
    /// <param name="lifetime"> LifeTime of Dependency Injection </param>
    public static void AddExcelWizardServices(this IServiceCollection services, bool isBlazorApp = false, ServiceLifetime lifetime = ServiceLifetime.Scoped)
    {
        if (isBlazorApp)
            services.AddBlazorDownloadFile(lifetime);

        if (lifetime == ServiceLifetime.Scoped)
        {
            services.AddScoped<IExcelWizardService, ExcelWizardService>();

            if (isBlazorApp is false)
                services.AddScoped<IBlazorDownloadFileService, FakeBlazorDownloadFileService>();
        }

        else if (lifetime == ServiceLifetime.Transient)
        {
            services.AddTransient<IExcelWizardService, ExcelWizardService>();

            if (isBlazorApp is false)
                services.AddTransient<IBlazorDownloadFileService, FakeBlazorDownloadFileService>();
        }

        else if (lifetime == ServiceLifetime.Singleton)
        {
            services.AddSingleton<IExcelWizardService, ExcelWizardService>();

            if (isBlazorApp is false)
                services.AddSingleton<IBlazorDownloadFileService, FakeBlazorDownloadFileService>();
        }

        else
        {
            throw new InvalidOperationException("The lifeTime is invalid");
        }
    }

    /// <summary>
    ///  Get Cell Column Number from Cell Column Letter, e.g. "A" => 1 or "C" => 3
    /// </summary>
    public static int GetCellColumnNumberByCellColumnLetter(this string cellColumnLetter)
    {
        if (string.IsNullOrWhiteSpace(cellColumnLetter))
            throw new ArgumentNullException(nameof(cellColumnLetter));

        int retVal = 0;
        string col = cellColumnLetter.ToUpper();
        for (int iChar = col.Length - 1; iChar >= 0; iChar--)
        {
            char colPiece = col[iChar];
            int colNum = colPiece - 64;
            retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
        }
        return retVal;
    }

    /// <summary>
    /// Get Cell Column Letter By Cell Column Number, e.g. 1 => "A" or 3 => "C"
    /// </summary>
    /// <param name="cellColumnNumber"></param>
    /// <returns></returns>
    public static string GetCellColumnLetterByCellColumnNumber(this int cellColumnNumber)
    {
        int dividend = cellColumnNumber;

        string cellName = string.Empty;

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            cellName = Convert.ToChar(65 + modulo) + cellName;
            dividend = (dividend - modulo) / 26;
        }

        return cellName.ToUpper();
    }
}