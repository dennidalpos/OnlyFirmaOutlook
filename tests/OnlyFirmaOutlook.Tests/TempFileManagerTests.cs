using OnlyFirmaOutlook.Services;

namespace OnlyFirmaOutlook.Tests;

public class TempFileManagerTests
{
    [Theory]
    [InlineData(@"\\server\share\file.docx", true)]
    [InlineData(@"C:\temp\file.docx", false)]
    public void IsUncPath_DetectsUncPaths(string path, bool expected)
    {
        var result = TempFileManager.IsUncPath(path);

        Assert.Equal(expected, result);
    }
}
