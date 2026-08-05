/*
 * OnlyFirmaOutlook
 * Copyright (c) 2026 Danny Perondi. All rights reserved.
 * Author: Danny Perondi
 * Proprietary and confidential.
 * Unauthorized copying, modification, distribution, sublicensing, disclosure,
 * or commercial use is prohibited without prior written permission.
 */

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

    public AssetProcessingResult ProcessImages(string html, string sourceHtmlPath, string assetsFolderPath)
    {
        Directory.CreateDirectory(assetsFolderPath);
        var assetsFolderName = Path.GetFileName(assetsFolderPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        var imgNodes = doc.DocumentNode.SelectNodes("//img[@src]");
        var vmlImageNodes = doc.DocumentNode
            .Descendants()
            .Where(node => HasNodeName(node, "v:imagedata") && node.Attributes["src"] != null)
            .ToList();
        var vmlNodes = doc.DocumentNode
            .Descendants()
            .Where(node => HasAnyAttribute(node, "o:href", "v:href", "xlink:href"))
            .ToList();
        var baseDir = Path.GetDirectoryName(sourceHtmlPath) ?? string.Empty;
        var pathMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (imgNodes != null)
        {
            foreach (var img in imgNodes)
            {
                ProcessAttribute(img, "src", baseDir, assetsFolderPath, assetsFolderName, pathMap);
            }
        }

        if (vmlImageNodes != null)
        {
            foreach (var vmlImage in vmlImageNodes)
            {
                ProcessAttribute(vmlImage, "src", baseDir, assetsFolderPath, assetsFolderName, pathMap);
            }
        }

        if (vmlNodes != null)
        {
            foreach (var node in vmlNodes)
            {
                ProcessAttribute(node, "o:href", baseDir, assetsFolderPath, assetsFolderName, pathMap);
                ProcessAttribute(node, "v:href", baseDir, assetsFolderPath, assetsFolderName, pathMap);
                ProcessAttribute(node, "xlink:href", baseDir, assetsFolderPath, assetsFolderName, pathMap);
            }
        }

        var processedHtml = doc.DocumentNode.InnerHtml;
        var plainText = SignatureInstaller.BuildPlainText(processedHtml);
        return new AssetProcessingResult(processedHtml, plainText);
    }

    private static bool HasAnyAttribute(HtmlNode node, params string[] attributeNames)
    {
        return attributeNames.Any(attributeName => node.Attributes[attributeName] != null);
    }

    private static bool HasNodeName(HtmlNode node, params string[] nodeNames)
    {
        return nodeNames.Any(nodeName => string.Equals(node.Name, nodeName, StringComparison.OrdinalIgnoreCase));
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

    private void ProcessAttribute(
        HtmlNode node,
        string attributeName,
        string baseDir,
        string assetsFolderPath,
        string assetsFolderName,
        Dictionary<string, string> pathMap)
    {
        var srcValue = node.GetAttributeValue(attributeName, string.Empty);
        if (string.IsNullOrWhiteSpace(srcValue))
        {
            return;
        }

        if (srcValue.StartsWith("cid:", StringComparison.OrdinalIgnoreCase) ||
            srcValue.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            srcValue.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        if (srcValue.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            if (TrySaveEmbeddedImage(srcValue, assetsFolderPath, assetsFolderName, out var rewrittenPath))
            {
                node.SetAttributeValue(attributeName, rewrittenPath);
            }
            return;
        }

        var resolvedPath = ResolveImagePath(srcValue, baseDir);
        if (resolvedPath == null || !File.Exists(resolvedPath))
        {
            _logger.LogWarning($"Immagine non trovata: {srcValue}");
            return;
        }

        if (!pathMap.TryGetValue(resolvedPath!, out var fileName))
        {
            if (IsFileInAssetsFolder(resolvedPath, assetsFolderPath))
            {
                fileName = Path.GetFileName(resolvedPath);
            }
            else
            {
                fileName = CreateStableFileName(resolvedPath);
                var destinationPath = Path.Combine(assetsFolderPath, fileName);
                if (!File.Exists(destinationPath))
                {
                    File.Copy(resolvedPath, destinationPath, overwrite: false);
                }
            }

            pathMap[resolvedPath] = fileName;
        }

        node.SetAttributeValue(attributeName, $"{assetsFolderName}/{fileName}");
    }

    private static bool IsFileInAssetsFolder(string filePath, string assetsFolderPath)
    {
        var normalizedFilePath = Path.GetFullPath(filePath);
        var normalizedAssetsFolder = Path.GetFullPath(assetsFolderPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;

        return normalizedFilePath.StartsWith(normalizedAssetsFolder, StringComparison.OrdinalIgnoreCase);
    }

    private bool TrySaveEmbeddedImage(
        string srcValue,
        string assetsFolderPath,
        string assetsFolderName,
        out string rewrittenPath)
    {
        rewrittenPath = string.Empty;

        try
        {
            const string prefix = "data:";
            var base64Index = srcValue.IndexOf("base64,", StringComparison.OrdinalIgnoreCase);
            if (base64Index <= 0)
            {
                _logger.LogWarning("Data URI non supportato (base64 mancante).");
                return false;
            }

            var meta = srcValue[prefix.Length..base64Index].TrimEnd(';');
            var mimeType = meta.Split(';')[0].Trim();
            var base64Data = srcValue[(base64Index + "base64,".Length)..].Trim();
            var bytes = Convert.FromBase64String(base64Data);
            var extension = GetExtensionFromMime(mimeType);
            var fileName = CreateStableFileName(bytes, extension);
            var destinationPath = Path.Combine(assetsFolderPath, fileName);

            if (!File.Exists(destinationPath))
            {
                File.WriteAllBytes(destinationPath, bytes);
            }

            rewrittenPath = $"{assetsFolderName}/{fileName}";
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning($"Impossibile salvare immagine embedded: {ex.Message}");
            return false;
        }
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

    private static string CreateStableFileName(byte[] content, string extension)
    {
        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(content);
        var hashString = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        var normalizedExtension = string.IsNullOrWhiteSpace(extension) ? ".img" : extension;
        return $"{hashString}{normalizedExtension}";
    }

    private static string GetExtensionFromMime(string mimeType)
    {
        return mimeType.ToLowerInvariant() switch
        {
            "image/png" => ".png",
            "image/jpeg" => ".jpg",
            "image/jpg" => ".jpg",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "image/svg+xml" => ".svg",
            "image/webp" => ".webp",
            "image/x-icon" => ".ico",
            "image/vnd.microsoft.icon" => ".ico",
            _ => ".img"
        };
    }
}

public record AssetProcessingResult(string Html, string PlainText);
