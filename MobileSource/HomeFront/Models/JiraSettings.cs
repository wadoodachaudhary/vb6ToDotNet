namespace HomeFront.Models;

/// <summary>
/// Connection settings for direct Jira import (entered once via the Jira settings dialog, saved
/// server-side under App_Data — NOT wwwroot, NOT the repo/deploy zip). The ApiToken is a secret.
/// Auth: if <see cref="Email"/> is set it's Jira Cloud Basic auth (email:token); if blank, the
/// token is used as a Bearer Personal Access Token (Jira Server / Data Center).
/// </summary>
public class JiraSettings
{
    public string BaseUrl { get; set; } = "";      // e.g. https://yourcompany.atlassian.net
    public string Email { get; set; } = "";         // Cloud account email (blank ⇒ Bearer/PAT)
    public string ApiToken { get; set; } = "";       // Cloud API token OR Server/DC PAT — SECRET
    public string ProjectKey { get; set; } = "";     // e.g. HF
    public string IssueType { get; set; } = "Bug";

    // Agile board that backs the project (rest/agile/1.0/board/{BoardId}/sprint). 0 ⇒ sprint
    // features (dropdown + push) are skipped. For HHM this is board 783.
    public int BoardId { get; set; } = 0;
    // Custom field that stores Sprint on an issue (company-managed Jira). For HHM: customfield_10007.
    public string SprintCustomFieldId { get; set; } = "customfield_10007";

    // Last successful synchronization timestamp (in UTC)
    public DateTime? LastSyncTime { get; set; }

    public bool IsConfigured =>
        !string.IsNullOrWhiteSpace(BaseUrl) &&
        !string.IsNullOrWhiteSpace(ProjectKey) &&
        !string.IsNullOrWhiteSpace(ApiToken);
}
