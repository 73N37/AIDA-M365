using AIDA.M365.Components;
using AIDA.M365.Models;
using AIDA.M365.Services;
using Bunit;
using FluentAssertions;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Moq;
using MudBlazor;
using MudBlazor.Services;
using Xunit;
using System.Collections.Generic;

namespace AIDA.M365.Tests.Components;

public class KanbanBoardTests : TestContext
{
    public KanbanBoardTests()
    {
        // Add MudBlazor services
        Services.AddMudServices();
        
        // Mock dependencies
        var mockService = new Mock<IOutlookCalendarEventService>();
        Services.AddSingleton(mockService.Object);
        
        var mockLogger = new Mock<ILogger<KanbanBoard>>();
        Services.AddSingleton(mockLogger.Object);

        // Mock MudBlazor JS calls
        JSInterop.SetupVoid("mudDragAndDrop.initDropZone", _ => true);
        JSInterop.SetupVoid("mudDragAndDrop.initContainer", _ => true);
    }

    [Fact]
    public void KanbanBoard_ShouldRenderSections()
    {
        // Arrange
        var sections = new List<KanbanSectionDefinition>
        {
            new() { Key = "todo", Title = "To Do", DayOffsetFromTodayUtc = 0, StartHourUtc = 9 },
            new() { Key = "done", Title = "Done", DayOffsetFromTodayUtc = 1, StartHourUtc = 9 }
        };
        var cards = new List<KanbanEventCard>();

        // Act
        var cut = RenderComponent<KanbanBoard>(parameters => parameters
            .Add(p => p.Sections, sections)
            .Add(p => p.Cards, cards)
        );

        // Assert
        cut.FindAll(".mud-paper").Count.Should().BeGreaterThan(0);
        foreach (var section in sections)
        {
            cut.Markup.Should().Contain(section.Title);
        }
    }
}
