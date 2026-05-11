using AIDA.M365.Services;
using Microsoft.Extensions.DependencyInjection;

namespace AIDA.M365.Extensions;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddAidaKanbanServices(this IServiceCollection services)
    {
        services.AddScoped<IOutlookCalendarEventService, OutlookCalendarEventService>();

        return services;
    }
}
