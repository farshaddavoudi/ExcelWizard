using System;
using BlazorDownloadFile;
using EasyExcelGenerator.Service;
using Microsoft.Extensions.DependencyInjection;

namespace EasyExcelGenerator;

public static class DIExtension
{
    /// <summary>
    /// Add EasyExcelGenerator package required services to IServiceCollection
    /// </summary>
    /// <param name="services"></param>
    /// <param name="lifetime"> LifeTime of Dependency Injection </param>
    public static void AddEasyExcelServices(this IServiceCollection services, ServiceLifetime lifetime = ServiceLifetime.Scoped)
    {
        services.AddBlazorDownloadFile(lifetime);

        if (lifetime == ServiceLifetime.Scoped)
            services.AddScoped<IEasyExcelService, EasyExcelService>();

        else if (lifetime == ServiceLifetime.Transient)
            services.AddTransient<IEasyExcelService, EasyExcelService>();

        else if (lifetime == ServiceLifetime.Singleton)
            services.AddSingleton<IEasyExcelService, EasyExcelService>();

        else
        {
            throw new InvalidOperationException("The lifeTime is invalid");
        }
    }
}