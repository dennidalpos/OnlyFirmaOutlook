using System.Security.Cryptography;
using HtmlAgilityPack;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace OnlyFirmaOutlook.Services;

public class AssetManager
{
    private readonly LoggingService _logger;

    public AssetManager()
    {
        _logger = LoggingService.Instance;
    }

    public AssetProcessingResult ProcessImages(string html, string sourceHtmlPath, string assetsFolderPath, string signatureName, bool useAbsolutePaths)
    {
        Directory.CreateDirectory(assetsFolderPath);
        var assetsFolderName = Path.GetFileName(assetsFolderPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        var imgNodes = doc.DocumentNode.SelectNodes("//img[@src]");
        var baseDir = Path.GetDirectoryName(sourceHtmlPath) ?? string.Empty;
        var pathMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (imgNodes != null)
        {
            foreach (var img in imgNodes)
            {
                var srcValue = img.GetAttributeValue("src", string.Empty);
                if (string.IsNullOrWhiteSpace(srcValue))
                {
                    continue;
                }

                if (srcValue.StartsWith("cid:", StringComparison.OrdinalIgnoreCase) ||
                    srcValue.StartsWith("data:", StringComparison.OrdinalIgnoreCase) ||
                    srcValue.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                    srcValue.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var resolvedPath = ResolveImagePath(srcValue, baseDir);
                if (resolvedPath == null || !File.Exists(resolvedPath))
                {
                    _logger.LogWarning($"Immagine non trovata: {srcValue}");
                    continue;
                }

                if (!pathMap.TryGetValue(resolvedPath, out var fileName))
                {
                    fileName = CreateStableFileName(resolvedPath);
                    var destinationPath = Path.Combine(assetsFolderPath, fileName);
                    if (!File.Exists(destinationPath))
                    {
                        File.Copy(resolvedPath, destinationPath, overwrite: false);
                    }
                    pathMap[resolvedPath] = fileName;
                }

                var rewritten = useAbsolutePaths
                    ? Path.Combine(assetsFolderPath, fileName)
                    : $"{assetsFolderName}/{fileName}";

                img.SetAttributeValue("src", rewritten);
            }
        }

        var processedHtml = doc.DocumentNode.InnerHtml;
        var plainText = SignatureInstaller.BuildPlainText(processedHtml);
        return new AssetProcessingResult(processedHtml, plainText);
    }

    private static string? ResolveImagePath(string srcValue, string baseDir)
    {
        if (srcValue.StartsWith("file:", StringComparison.OrdinalIgnoreCase) &&
            Uri.TryCreate(srcValue, UriKind.Absolute, out var uri) &&
            uri.IsFile)
        {
            return uri.LocalPath;
        }

        if (Path.IsPathRooted(srcValue))
        {
            return srcValue;
        }

        var combined = Path.Combine(baseDir, srcValue);
        return combined;
    }

    private static string CreateStableFileName(string sourcePath)
    {
        using var sha = SHA256.Create();
        var bytes = File.ReadAllBytes(sourcePath);
        var hash = sha.ComputeHash(bytes);
        var hashString = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        var extension = Path.GetExtension(sourcePath);
        if (string.IsNullOrWhiteSpace(extension))
        {
            extension = ".img";
        }

        return $"{hashString}{extension}";
    }
}

public record AssetProcessingResult(string Html, string PlainText);
