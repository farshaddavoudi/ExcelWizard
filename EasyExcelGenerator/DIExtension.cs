using EasyExcelGenerator.Service;
using Microsoft.Extensions.DependencyInjection;

namespace EasyExcelGenerator;

public static class DIExtension
{
    /// <summary>
    /// Add EasyExcelGenerator package required services to IServiceCollection
    /// </summary>
    /// <param name="services"></param>
    public static void AddEasyExcelServices(this IServiceCollection services)
    {
        services.AddScoped<IEasyExcelService, EasyExcelService>();
    }
}