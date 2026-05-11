using AIDA.M365.Models;
using Microsoft.Extensions.Logging;

namespace AIDA.M365.Services;

public sealed class EventCommandCenterService : IEventCommandCenterService
{
    private readonly IOutlookCalendarEventService _outlookCalendarEventService;
    private readonly IDynamicsEventLinkService _dynamicsEventLinkService;
    private readonly ICardCustomizationRepository _cardCustomizationRepository;
    private readonly IAiCardSummaryService _aiCardSummaryService;
    private readonly ILogger<EventCommandCenterService> _logger;

    public EventCommandCenterService(
        IOutlookCalendarEventService outlookCalendarEventService,
        IDynamicsEventLinkService dynamicsEventLinkService,
        ICardCustomizationRepository cardCustomizationRepository,
        IAiCardSummaryService aiCardSummaryService,
        ILogger<EventCommandCenterService> logger)
    {
        _outlookCalendarEventService = outlookCalendarEventService;
        _dynamicsEventLinkService = dynamicsEventLinkService;
        _cardCustomizationRepository = cardCustomizationRepository;
        _aiCardSummaryService = aiCardSummaryService;
        _logger = logger;
    }

    public async Task<MoveCardResult> MoveCardAsync(
        MoveCardRequest request,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(request.Card.GraphEventId))
        {
            return MoveCardResult.Failure("Graph event id is required.");
        }

        if (request.EndUtc <= request.StartUtc)
        {
            return MoveCardResult.Failure("End must be greater than start.");
        }

        try
        {
            await _outlookCalendarEventService.UpdateEventTimeAsync(
                request.Card.GraphEventId,
                request.StartUtc,
                request.EndUtc,
                cancellationToken);

            if (request.Card.IsLinkedToDynamics)
            {
                await _dynamicsEventLinkService.UpdateLinkedEventAsync(
                    new DynamicsEventUpdate(
                        request.Card.DynamicsEntityLogicalName!,
                        request.Card.DynamicsEntityId!,
                        request.Card.GraphEventId,
                        request.TargetSectionKey,
                        request.StartUtc,
                        request.EndUtc),
                    cancellationToken);
            }

            await _cardCustomizationRepository.SaveAsync(
                new CardCustomization(
                    request.Card.MetadataId,
                    request.Card.GraphEventId,
                    request.TargetSectionKey,
                    request.StartUtc,
                    request.EndUtc,
                    request.Card.ColorHex,
                    request.Card.Priority,
                    request.Card.SmartSummary,
                    request.Card.ActionItems,
                    DateTimeOffset.UtcNow),
                cancellationToken);

            return MoveCardResult.Success();
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(
                ex,
                "Failed to move event {EventId} to section {SectionKey}",
                request.Card.GraphEventId,
                request.TargetSectionKey);

            return MoveCardResult.Failure("The event could not be synchronized across Microsoft 365.");
        }
    }

    public Task<CardSummaryResult> SummarizeCardAsync(
        KanbanEventCard card,
        CancellationToken cancellationToken = default) =>
        _aiCardSummaryService.SummarizeAsync(card, cancellationToken);
}
