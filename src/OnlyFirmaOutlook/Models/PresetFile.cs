namespace OnlyFirmaOutlook.Models;




public class PresetFile
{
    public string FullPath { get; set; } = string.Empty;
    public string FileName { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;

    public override string ToString() => DisplayName;
}
