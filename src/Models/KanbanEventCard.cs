using System;

namespace AIDA.M365.Models;

public sealed class KanbanEventCard
{
    public required string GraphEventId { get; init; }

    public required string Subject { get; set; }

    public required string SectionKey { get; set; }

    // Store times in UTC to keep drag-drop updates deterministic.
    public required DateTimeOffset StartUtc { get; set; }

    public required DateTimeOffset EndUtc { get; set; }

    public TimeSpan Duration => EndUtc - StartUtc;
}
