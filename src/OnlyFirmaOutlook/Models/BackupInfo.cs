namespace OnlyFirmaOutlook.Models;

public class BackupInfo
{
    public string FileName { get; set; } = string.Empty;
    public string FullPath { get; set; } = string.Empty;
    public DateTime CreatedAt { get; set; }
    public long SizeBytes { get; set; }

    public string DisplayInfo => $"{CreatedAt:dd/MM/yyyy HH:mm} - {FileName} ({FormatSize(SizeBytes)})";

    private static string FormatSize(long bytes)
    {
        string[] sizes = { "B", "KB", "MB", "GB" };
        var len = (double)bytes;
        var order = 0;

        while (len >= 1024 && order < sizes.Length - 1)
        {
            order++;
            len /= 1024;
        }

        return $"{len:0.##} {sizes[order]}";
    }
}
