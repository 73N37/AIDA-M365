using System;
using System.Collections.Generic;

namespace AIDA.M365.Models;

public sealed class KanbanEventCard
{
    public required string GraphEventId { get; init; }

    public string? DynamicsEntityId { get; init; }

    public string? DynamicsEntityLogicalName { get; init; }

    public Guid? MetadataId { get; init; }

    public required string Subject { get; set; }

    public string? BodyContent { get; set; }

    public required string SectionKey { get; set; }

    // Store times in UTC to keep drag-drop updates deterministic.
    public required DateTimeOffset StartUtc { get; set; }

    public required DateTimeOffset EndUtc { get; set; }

    public string? ColorHex { get; set; }

    public string? Priority { get; set; }

    public string? SmartSummary { get; set; }

    public IReadOnlyList<string> ActionItems { get; set; } = [];

    public TimeSpan Duration => EndUtc - StartUtc;

    public bool IsLinkedToDynamics =>
        !string.IsNullOrWhiteSpace(DynamicsEntityId)
        && !string.IsNullOrWhiteSpace(DynamicsEntityLogicalName);
}
