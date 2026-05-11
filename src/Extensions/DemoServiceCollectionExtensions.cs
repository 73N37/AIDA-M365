using AIDA.M365.Services;
using Microsoft.Extensions.DependencyInjection;

namespace AIDA.M365.Extensions;

public static class DemoServiceCollectionExtensions
{
    public static IServiceCollection AddAidaKanbanDemoServices(this IServiceCollection services)
    {
        services.AddScoped<IOutlookCalendarEventService, DemoOutlookCalendarEventService>();
        services.AddScoped<IEventCommandCenterService, DemoEventCommandCenterService>();

        return services;
    }
}
