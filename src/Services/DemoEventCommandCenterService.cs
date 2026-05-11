using AIDA.M365.Models;
using Microsoft.Extensions.Logging;

namespace AIDA.M365.Services;

public sealed class DemoEventCommandCenterService : IEventCommandCenterService
{
    private readonly IOutlookCalendarEventService _outlookCalendarEventService;
    private readonly ILogger<DemoEventCommandCenterService> _logger;

    public DemoEventCommandCenterService(
        IOutlookCalendarEventService outlookCalendarEventService,
        ILogger<DemoEventCommandCenterService> logger)
    {
        _outlookCalendarEventService = outlookCalendarEventService;
        _logger = logger;
    }

    public async Task<MoveCardResult> MoveCardAsync(
        MoveCardRequest request,
        CancellationToken cancellationToken = default)
    {
        await _outlookCalendarEventService.UpdateEventTimeAsync(
            request.Card.GraphEventId,
            request.StartUtc,
            request.EndUtc,
            cancellationToken);

        if (request.Card.IsLinkedToDynamics)
        {
            _logger.LogInformation(
                "[DEMO] Simulated Dynamics update for {EntityLogicalName} {EntityId}",
                request.Card.DynamicsEntityLogicalName,
                request.Card.DynamicsEntityId);
        }

        _logger.LogInformation(
            "[DEMO] Simulated metadata save for event {EventId}",
            request.Card.GraphEventId);

        return MoveCardResult.Success();
    }

    public Task<CardSummaryResult> SummarizeCardAsync(
        KanbanEventCard card,
        CancellationToken cancellationToken = default)
    {
        var summary = $"Summary for {card.Subject}: review the event details and confirm next steps.";
        string[] actions =
        [
            "Confirm owner and due date",
            "Update related Dynamics record",
            "Send follow-up notes to attendees"
        ];

        return Task.FromResult(CardSummaryResult.Success(summary, actions));
    }
}
