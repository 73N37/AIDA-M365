using AIDA.M365.Models;
using AIDA.M365.Services;
using FluentAssertions;
using Microsoft.Extensions.Logging;
using Moq;
using Xunit;

namespace AIDA.M365.Tests.Services;

public sealed class EventCommandCenterServiceTests
{
    private readonly Mock<IOutlookCalendarEventService> _outlook = new();
    private readonly Mock<IDynamicsEventLinkService> _dynamics = new();
    private readonly Mock<ICardCustomizationRepository> _repository = new();
    private readonly Mock<IAiCardSummaryService> _ai = new();
    private readonly Mock<ILogger<EventCommandCenterService>> _logger = new();

    [Fact]
    public async Task MoveCardAsync_ShouldPatchGraphAndSkipDynamics_WhenCardIsNotLinked()
    {
        var service = CreateService();
        var card = CreateCard();
        var start = DateTimeOffset.UtcNow;
        var end = start.AddHours(1);

        var result = await service.MoveCardAsync(new MoveCardRequest(card, "planned", start, end));

        result.Succeeded.Should().BeTrue();
        _outlook.Verify(x => x.UpdateEventTimeAsync(card.GraphEventId, start, end, It.IsAny<CancellationToken>()), Times.Once);
        _dynamics.Verify(x => x.UpdateLinkedEventAsync(It.IsAny<DynamicsEventUpdate>(), It.IsAny<CancellationToken>()), Times.Never);
        _repository.Verify(x => x.SaveAsync(It.Is<CardCustomization>(c => c.GraphEventId == card.GraphEventId), It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task MoveCardAsync_ShouldUpdateDynamics_WhenCardIsLinked()
    {
        var service = CreateService();
        var card = CreateCard().WithDynamics();
        var start = DateTimeOffset.UtcNow;
        var end = start.AddHours(1);

        var result = await service.MoveCardAsync(new MoveCardRequest(card, "planned", start, end));

        result.Succeeded.Should().BeTrue();
        _dynamics.Verify(
            x => x.UpdateLinkedEventAsync(
                It.Is<DynamicsEventUpdate>(u => u.EntityId == card.DynamicsEntityId && u.EntityLogicalName == card.DynamicsEntityLogicalName),
                It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task MoveCardAsync_ShouldReturnFailure_WhenEndIsBeforeStart()
    {
        var service = CreateService();
        var start = DateTimeOffset.UtcNow;

        var result = await service.MoveCardAsync(new MoveCardRequest(CreateCard(), "planned", start, start.AddMinutes(-1)));

        result.Succeeded.Should().BeFalse();
        _outlook.Verify(x => x.UpdateEventTimeAsync(It.IsAny<string>(), It.IsAny<DateTimeOffset>(), It.IsAny<DateTimeOffset>(), It.IsAny<CancellationToken>()), Times.Never);
    }

    private EventCommandCenterService CreateService() =>
        new(_outlook.Object, _dynamics.Object, _repository.Object, _ai.Object, _logger.Object);

    private static KanbanEventCard CreateCard() =>
        new()
        {
            GraphEventId = "event-1",
            Subject = "Test",
            SectionKey = "backlog",
            StartUtc = DateTimeOffset.UtcNow,
            EndUtc = DateTimeOffset.UtcNow.AddHours(1)
        };
}

file static class KanbanEventCardTestExtensions
{
    public static KanbanEventCard WithDynamics(this KanbanEventCard card)
    {
        return new KanbanEventCard
        {
            GraphEventId = card.GraphEventId,
            DynamicsEntityId = "account-id",
            DynamicsEntityLogicalName = "account",
            Subject = card.Subject,
            SectionKey = card.SectionKey,
            StartUtc = card.StartUtc,
            EndUtc = card.EndUtc
        };
    }
}
