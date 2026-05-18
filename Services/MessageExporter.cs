namespace TeamsMessageFetcher;

// ─── CSV Exporter ─────────────────────────────────────────────────────────────

public class MessageExporter
{
    public void ExportAll(
        List<TeamMessage>  messages,
        List<CalendarEvent> events,
        List<MailMessage>  emails,
        List<DevOpsCommit> commits,
        DateOnly date)
    {
        var fileName = $"daily_summary_{date:yyyy-MM-dd}.csv";
        var path     = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);

        var lines = new List<string>();

        AppendMessages(lines, messages);
        AppendEvents(lines, events);
        AppendMail(lines, emails);
        AppendCommits(lines, commits);

        File.WriteAllLines(path, lines, System.Text.Encoding.UTF8);
        Console.WriteLine($"\n💾 Exported to: {path}");
    }

    // ─── Section writers ──────────────────────────────────────────────────────

    private static void AppendMessages(List<string> lines, List<TeamMessage> messages)
    {
        lines.Add("=== TEAMS MESSAGES ===");
        lines.Add("Timestamp,Source,Sender,SentByMe,Preview");

        foreach (var msg in messages.OrderBy(m => m.Timestamp))
        {
            lines.Add(Row(
                msg.Timestamp.ToString("HH:mm:ss"),
                msg.Source,
                msg.SenderName,
                msg.IsSentByMe ? "Yes" : "No",
                msg.BodyPreview));
        }

        lines.Add(string.Empty);
    }

    private static void AppendEvents(List<string> lines, List<CalendarEvent> events)
    {
        lines.Add("=== CALENDAR EVENTS ===");
        lines.Add("Start,End,Duration,Subject,Organizer,IsOrganizer,IsOnline,Attendees");

        foreach (var ev in events.OrderBy(e => e.Start))
        {
            lines.Add(Row(
                ev.Start.ToString("HH:mm"),
                ev.End.ToString("HH:mm"),
                $"{(int)ev.Duration.TotalMinutes} min",
                ev.Subject,
                ev.Organizer,
                ev.IsOrganizer ? "Yes" : "No",
                ev.IsOnline ? "Yes" : "No",
                string.Join("; ", ev.Attendees)));
        }

        lines.Add(string.Empty);
    }

    private static void AppendMail(List<string> lines, List<MailMessage> emails)
    {
        lines.Add("=== EMAILS ===");
        lines.Add("Timestamp,Direction,From,To,Subject,Preview");

        foreach (var mail in emails.OrderBy(m => m.Timestamp))
        {
            lines.Add(Row(
                mail.Timestamp.ToString("HH:mm:ss"),
                mail.IsSent ? "Sent" : "Received",
                mail.From,
                string.Join("; ", mail.To),
                mail.Subject,
                mail.BodyPreview));
        }

        lines.Add(string.Empty);
    }

    private static void AppendCommits(List<string> lines, List<DevOpsCommit> commits)
    {
        lines.Add("=== AZURE DEVOPS COMMITS ===");
        lines.Add("Timestamp,Project,Repository,CommitId,Message,Url");

        foreach (var commit in commits.OrderBy(c => c.Timestamp))
        {
            lines.Add(Row(
                commit.Timestamp.ToString("HH:mm:ss"),
                commit.Project,
                commit.Repository,
                commit.CommitId[..7],
                commit.Message,
                commit.Url));
        }

        lines.Add(string.Empty);
    }

    // ─── CSV helpers ──────────────────────────────────────────────────────────

    private static string Row(params string[] cols) =>
        string.Join(",", cols.Select(Escape));

    private static string Escape(string value)
    {
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n'))
            return $"\"{value.Replace("\"", "\"\"")}\"";
        return value;
    }
}
