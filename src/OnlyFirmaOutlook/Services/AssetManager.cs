using System.IO;
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
        var vmlNodes = doc.DocumentNode.SelectNodes("//*[@o:href or @v:href or @xlink:href]");
        var vmlImageNodes = doc.DocumentNode.Descendants()
            .Where(node => node.Name.Contains("imagedata", StringComparison.OrdinalIgnoreCase)
                           && node.Attributes["src"] != null)
            .ToList();
        var baseDir = Path.GetDirectoryName(sourceHtmlPath) ?? string.Empty;
        var assetSearchFolders = BuildAssetSearchFolders(baseDir);
        var pathMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (imgNodes != null)
        {
            foreach (var img in imgNodes)
            {
                ProcessAttribute(img, "src", baseDir, assetSearchFolders, assetsFolderPath, assetsFolderName, useAbsolutePaths, pathMap);
            }
        }

        if (vmlNodes != null)
        {
            foreach (var node in vmlNodes)
            {
                ProcessAttribute(node, "o:href", baseDir, assetSearchFolders, assetsFolderPath, assetsFolderName, useAbsolutePaths, pathMap);
                ProcessAttribute(node, "v:href", baseDir, assetSearchFolders, assetsFolderPath, assetsFolderName, useAbsolutePaths, pathMap);
                ProcessAttribute(node, "xlink:href", baseDir, assetSearchFolders, assetsFolderPath, assetsFolderName, useAbsolutePaths, pathMap);
            }
        }

        if (vmlImageNodes.Count > 0)
        {
            foreach (var node in vmlImageNodes)
            {
                ProcessAttribute(node, "src", baseDir, assetSearchFolders, assetsFolderPath, assetsFolderName, useAbsolutePaths, pathMap);
            }
        }

        var processedHtml = doc.DocumentNode.InnerHtml;
        var plainText = SignatureInstaller.BuildPlainText(processedHtml);
        return new AssetProcessingResult(processedHtml, plainText);
    }

    private static string? ResolveImagePath(string srcValue, string baseDir, IReadOnlyList<string> assetSearchFolders)
    {
        var normalized = Uri.UnescapeDataString(srcValue.Trim());

        if (normalized.StartsWith("file:", StringComparison.OrdinalIgnoreCase) &&
            Uri.TryCreate(normalized, UriKind.Absolute, out var normalizedUri) &&
            normalizedUri.IsFile)
        {
            return normalizedUri.LocalPath;
        }

        if (srcValue.StartsWith("file:", StringComparison.OrdinalIgnoreCase) &&
            Uri.TryCreate(srcValue, UriKind.Absolute, out var uri) &&
            uri.IsFile)
        {
            return uri.LocalPath;
        }

        if (Path.IsPathRooted(normalized))
        {
            return normalized;
        }

        var combined = Path.Combine(baseDir, normalized);
        if (File.Exists(combined))
        {
            return combined;
        }

        foreach (var folder in assetSearchFolders)
        {
            var candidate = Path.Combine(folder, normalized);
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        var fileName = Path.GetFileName(normalized);
        if (!string.IsNullOrWhiteSpace(fileName) && !fileName.Equals(normalized, StringComparison.OrdinalIgnoreCase))
        {
            foreach (var folder in assetSearchFolders)
            {
                var candidate = Path.Combine(folder, fileName);
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
        }

        return combined;
    }

    private void ProcessAttribute(
        HtmlNode node,
        string attributeName,
        string baseDir,
        IReadOnlyList<string> assetSearchFolders,
        string assetsFolderPath,
        string assetsFolderName,
        bool useAbsolutePaths,
        Dictionary<string, string> pathMap)
    {
        var srcValue = node.GetAttributeValue(attributeName, string.Empty);
        if (string.IsNullOrWhiteSpace(srcValue))
        {
            return;
        }

        if (srcValue.StartsWith("cid:", StringComparison.OrdinalIgnoreCase) ||
            srcValue.StartsWith("data:", StringComparison.OrdinalIgnoreCase) ||
            srcValue.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            srcValue.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var resolvedPath = ResolveImagePath(srcValue, baseDir, assetSearchFolders);
        if (resolvedPath == null || !File.Exists(resolvedPath))
        {
            _logger.LogWarning($"Immagine non trovata: {srcValue}");
            return;
        }

        if (!pathMap.TryGetValue(resolvedPath!, out var fileName))
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

        node.SetAttributeValue(attributeName, rewritten);
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

    private static IReadOnlyList<string> BuildAssetSearchFolders(string baseDir)
    {
        var folders = new List<string> { baseDir };

        if (!Directory.Exists(baseDir))
        {
            return folders;
        }

        foreach (var folder in Directory.EnumerateDirectories(baseDir, "*_files"))
        {
            folders.Add(folder);
        }

        foreach (var folder in Directory.EnumerateDirectories(baseDir, "*_file"))
        {
            folders.Add(folder);
        }

        return folders;
    }
}

public record AssetProcessingResult(string Html, string PlainText);
