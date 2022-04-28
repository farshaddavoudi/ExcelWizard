using BlazorDownloadFile;
using ExcelWizard.Service;
using Microsoft.Extensions.DependencyInjection;
using System;

namespace ExcelWizard;

public static class DIExtension
{
    /// <summary>
    /// Add ExcelWizard package required services to IServiceCollection
    /// </summary>
    /// <param name="services"></param>
    /// <param name="isBlazorApp"> Do you register for a Blazor app or not. In case of API or MVC project, is will be false </param>
    /// <param name="lifetime"> LifeTime of Dependency Injection </param>
    public static void AddEasyExcelServices(this IServiceCollection services, bool isBlazorApp = false, ServiceLifetime lifetime = ServiceLifetime.Scoped)
    {
        if (isBlazorApp)
            services.AddBlazorDownloadFile(lifetime);

        if (lifetime == ServiceLifetime.Scoped)
        {
            services.AddScoped<IEasyExcelService, EasyExcelService>();

            if (isBlazorApp is false)
                services.AddScoped<IBlazorDownloadFileService, FakeBlazorDownloadFileService>();
        }

        else if (lifetime == ServiceLifetime.Transient)
        {
            services.AddTransient<IEasyExcelService, EasyExcelService>();

            if (isBlazorApp is false)
                services.AddTransient<IBlazorDownloadFileService, FakeBlazorDownloadFileService>();
        }

        else if (lifetime == ServiceLifetime.Singleton)
        {
            services.AddSingleton<IEasyExcelService, EasyExcelService>();

            if (isBlazorApp is false)
                services.AddSingleton<IBlazorDownloadFileService, FakeBlazorDownloadFileService>();
        }

        else
        {
            throw new InvalidOperationException("The lifeTime is invalid");
        }
    }
}