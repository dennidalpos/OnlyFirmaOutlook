namespace OnlyFirmaOutlook.Models;




public class SignatureInfo
{
    public string Name { get; set; } = string.Empty;
    public string FolderPath { get; set; } = string.Empty;
    public bool HasHtm { get; set; }
    public bool HasRtf { get; set; }
    public bool HasTxt { get; set; }
    public bool HasFilesFolder { get; set; }
    public bool HasFileFolder { get; set; }

    public string DisplayInfo
    {
        get
        {
            var parts = new List<string>();
            if (HasHtm) parts.Add("HTM");
            if (HasRtf) parts.Add("RTF");
            if (HasTxt) parts.Add("TXT");
            if (HasFilesFolder || HasFileFolder) parts.Add("Assets");

            return $"{Name} ({string.Join(", ", parts)})";
        }
    }

    public override string ToString() => DisplayInfo;
}
