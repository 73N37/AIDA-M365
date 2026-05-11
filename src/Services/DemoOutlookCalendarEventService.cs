using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace AIDA.M365.Services;

public sealed class DemoOutlookCalendarEventService : IOutlookCalendarEventService
{
    private readonly ILogger<DemoOutlookCalendarEventService> _logger;

    public DemoOutlookCalendarEventService(ILogger<DemoOutlookCalendarEventService> logger)
    {
        _logger = logger;
    }

    public Task UpdateEventTimeAsync(
        string graphEventId,
        DateTimeOffset newStartUtc,
        DateTimeOffset newEndUtc,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation(
            "[DEMO] Simulated Graph PATCH for event {EventId}. Start={StartUtc}, End={EndUtc}",
            graphEventId,
            newStartUtc,
            newEndUtc);

        return Task.CompletedTask;
    }
}
