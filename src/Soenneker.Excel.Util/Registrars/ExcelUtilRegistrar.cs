using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Soenneker.Excel.Util.Abstract;

namespace Soenneker.Excel.Util.Registrars;

/// <summary>
/// Provides methods for reading and writing Excel files using strongly-typed objects with automatic property mapping and basic type conversion
/// </summary>
public static class ExcelUtilRegistrar
{
    /// <summary>
    /// Adds <see cref="IExcelUtil"/> as a singleton service. <para/>
    /// </summary>
    public static IServiceCollection AddExcelUtilAsSingleton(this IServiceCollection services)
    {
        services.TryAddSingleton<IExcelUtil, ExcelUtil>();

        return services;
    }

    /// <summary>
    /// Adds <see cref="IExcelUtil"/> as a scoped service. <para/>
    /// </summary>
    public static IServiceCollection AddExcelUtilAsScoped(this IServiceCollection services)
    {
        services.TryAddScoped<IExcelUtil, ExcelUtil>();

        return services;
    }
}
