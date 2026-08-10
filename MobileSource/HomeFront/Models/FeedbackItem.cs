using System;
using System.Collections.Generic;

namespace HomeFront.Models;

/// <summary>
/// One feedback/test "ticket". Persisted under wwwroot/tickets as a JSON record plus a
/// companion annotated PNG and a printable HTML document. <see cref="TicketNumber"/> is the
/// unique key — the next number is max(existing)+1 (see FeedbackService.GetNextTicketNumberAsync).
/// </summary>
public class FeedbackItem
{
    // Unique key. Assigned on save as max(existing tickets)+1. Drives the ticket_NNNNN_* filenames.
    public int TicketNumber { get; set; }

    public string Id { get; set; } = string.Empty;
    public string Topic { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    // Jira wiki-markup form of the description, captured live from EditorLiteControl (ReadAsync →
    // JiraText) so "Export to Jira" sends proper formatting instead of a plain-text strip. Empty for
    // legacy plain-text tickets (export then falls back to an HTML→text conversion of Description).
    public string DescriptionJira { get; set; } = string.Empty;
    public string UserName { get; set; } = string.Empty;
    public DateTime CreatedDate { get; set; }

    /// <summary>UTC time of the last LOCAL save (stamped by UpdateFeedbackAsync). The Jira pull-sync
    /// compares this against the issue's board timestamp so a STALE search-index snapshot can never
    /// overwrite a fresher local edit (Jira's issue search lags real changes by minutes).</summary>
    public DateTime? LastLocalEditDate { get; set; }
    public string Url { get; set; } = string.Empty;

    // External tracker key when this ticket was imported from another system (e.g. Jira "HHM-12").
    // Blank for tickets created natively in the app — the board/ticket then shows the "PM-{number}" label.
    public string ExternalKey { get; set; } = string.Empty;

    // Jira-style workflow fields shown in the ticket screen's Details sidebar. File-serialised
    // only — no DB column. Status drives the To Do / In Progress / Done dropdown. Category is the
    // free-text project/area used to box tickets on the Feedback home board.
    public string Status { get; set; } = "To Do";
    // Closing reason captured in the "Close issue" dialog when the ticket is moved to Done.
    public string Resolution { get; set; } = string.Empty;
    public string Priority { get; set; } = "Medium";
    public string Category { get; set; } = string.Empty;
    public string Assignee { get; set; } = string.Empty;   // editable-dropdown person field (UserName = Reporter)

    // ── Jira bug-ticket fields ──
    public string IssueType { get; set; } = "Bug";          // Bug, Feature Request, Enhancement, Migration Gap, VB6 Incompatibility, Question
    public string Severity { get; set; } = "Major";         // Critical, Major, Minor, Trivial
    public string Labels { get; set; } = string.Empty;      // comma-separated (e.g. frontend, critical)
    public string FixVersion { get; set; } = string.Empty;  // fix version / milestone
    public DateTime? DueDate { get; set; }                  // deadline for the fix

    // ── Jira "Details" sidebar fields (parity with the real HHM board) ──
    public string Parent { get; set; } = string.Empty;        // parent epic / issue key
    public string Team { get; set; } = string.Empty;          // assigned team
    public DateTime? StartDate { get; set; }                  // work start date
    public string Sprint { get; set; } = string.Empty;        // sprint, e.g. "HYP Sprint 1"
    public string StoryPoints { get; set; } = string.Empty;   // story point estimate (free text; "" = None)

    // Structured bug-report body (legacy — folded into Description as bold sections; kept for back-compat /
    // the Jira export payload). No longer surfaced as separate editable fields in the ticket screen.
    public string EnvironmentInfo { get; set; } = string.Empty;   // OS / Browser / Device / App version
    public string Prerequisites { get; set; } = string.Empty;     // setup steps / initial state
    public string StepsToReproduce { get; set; } = string.Empty;
    public string ActualResult { get; set; } = string.Empty;      // what the user actually sees
    public string ExpectedResult { get; set; } = string.Empty;    // what should happen per spec

    // ── HomeFront VB6→Blazor migration context ──
    public string VB6Form { get; set; } = string.Empty;     // originating VB6 form, e.g. FJob.frm
    public string Module { get; set; } = string.Empty;      // subsystem: Estimating, Reports, Vendors, …

    // Context captured at F8 time: the screen/form the user was on + the focused control.
    public string ScreenName { get; set; } = string.Empty;
    public string FocusedControl { get; set; } = string.Empty;

    // Relative web url to the PRIMARY annotated PNG, e.g. /tickets/ticket_00007_jdoe_20260618-013055.png.
    // Drives the board card thumbnail. For multi-attachment tickets it is the first image in Attachments.
    public string ScreenshotPath { get; set; } = string.Empty;

    // All files attached to the ticket (images + any docs). Empty for legacy single-screenshot tickets,
    // which still render via ScreenshotPath. Populated for imports (e.g. a Jira issue with several images).
    public List<TicketAttachment> Attachments { get; set; } = new();

    // Relative web url to the printable HTML document for this ticket.
    public string DocumentPath { get; set; } = string.Empty;

    // The vector annotation objects (text/rect/arrow-text/rectangle-text/pen) as JSON, so a
    // ticket can be re-opened and re-edited later. The flattened PNG is the canonical image.
    public string AnnotationsJson { get; set; } = string.Empty;

    // Diagnostic context (active form properties/fields, session settings, SQL execution history, sequence of events)
    public string DiagnosticsJson { get; set; } = string.Empty;

    // Jira-style comment thread. File-serialised on the ticket JSON only — NOT included in the Jira export payload.
    public List<TicketComment> Comments { get; set; } = new();

    // ── Linked work items (HHM-216) — Jira-style issue links ──
    // Directional, from THIS ticket's perspective ("blocks HHM-12", "is blocked by HHM-9").
    // Pushed to the Jira board on save; refreshed from Jira's issuelinks on sync.
    public List<TicketLink> Links { get; set; } = new();
}

/// <summary>One linked work item (HHM-216). <see cref="Type"/> is the directional phrase from this
/// ticket's perspective (relates to / blocks / is blocked by / duplicates / is duplicated by /
/// causes / is caused by); <see cref="TargetKey"/> is the other ticket's display key ("HHM-116",
/// or "PM-12" for a local-only ticket); <see cref="TargetSummary"/> is cached for display when the
/// target isn't on the local board.</summary>
public class TicketLink
{
    public string Type { get; set; } = "relates to";
    public string TargetKey { get; set; } = string.Empty;
    public string TargetSummary { get; set; } = string.Empty;
}

/// <summary>One comment in a ticket's discussion thread.</summary>
public class TicketComment
{
    public string Id { get; set; } = string.Empty;   // stable id for edit/delete targeting
    public string Author { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public DateTime Date { get; set; }
    public DateTime? EditedDate { get; set; }         // set when the comment is edited
    /// <summary>Jira's own comment id once this comment exists on the board — set when we push it
    /// and when we pull one down. It is what keeps the two sides in step: a pull skips comments it
    /// already has, so our own pushed comment never comes back as a duplicate.</summary>
    public string? ExternalId { get; set; }
}

/// <summary>One activity/history entry for a ticket — a field change pulled from the Jira changelog
/// (HHM-147). Not persisted locally; fetched live for tickets that exist on the board.</summary>
public record TicketActivity(string Author, string Field, string OldValue, string NewValue, DateTime Date);

/// <summary>One file attached to a ticket. <see cref="Path"/> is the relative web url under
/// /tickets/...; <see cref="ContentType"/> distinguishes images (editable in the viewer) from other files.</summary>
public class TicketAttachment
{
    public string Path { get; set; } = string.Empty;          // web url, e.g. /tickets/jira-HHM/ticket_05052_..._a2.png
    public string FileName { get; set; } = string.Empty;      // original/display name
    public string ContentType { get; set; } = string.Empty;   // e.g. image/png
}
