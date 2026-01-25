using System.IO;
using System.Text.RegularExpressions;
using OnlyFirmaOutlook.Models;

namespace OnlyFirmaOutlook.Services;

public static class WordHtmlCidPostProcessor
{
    private static readonly Regex ImageSrcRegex = new(
        @"<\s*(?<tag>img|v:imagedata)\b[^>]*?\bsrc\s*=\s*(?<value>('(?<inner>[^']*)'|""(?<inner>[^""]*)""|(?<inner>[^\s>]+)))",
        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);

    public static (string Html, IReadOnlyList<InlineImage> Images) RewriteLocalImageRefsToCid(
        string html,
        string baseDirectory)
    {
        if (html is null)
        {
            throw new ArgumentNullException(nameof(html));
        }

        if (string.IsNullOrWhiteSpace(baseDirectory))
        {
            throw new ArgumentException("Base directory is required.", nameof(baseDirectory));
        }

        var normalizedBaseDirectory = Path.GetFullPath(baseDirectory);
        var imagesByPath = new Dictionary<string, InlineImage>(StringComparer.OrdinalIgnoreCase);

        string rewrittenHtml = ImageSrcRegex.Replace(html, match =>
        {
            var srcValue = match.Groups["inner"].Value;
            if (ShouldIgnoreSource(srcValue))
            {
                return match.Value;
            }

            var resolvedPath = ResolveLocalPath(srcValue, normalizedBaseDirectory);
            if (resolvedPath is null || !File.Exists(resolvedPath))
            {
                return match.Value;
            }

            if (!imagesByPath.TryGetValue(resolvedPath!, out var inlineImage))
            {
                var fileName = Path.GetFileName(resolvedPath);
                var contentId = $"{Guid.NewGuid():N}@onlyfirmaoutlook";
                inlineImage = new InlineImage(contentId, resolvedPath, fileName);
                imagesByPath[resolvedPath] = inlineImage;
            }

            var innerGroup = match.Groups["inner"];
            var relativeStart = innerGroup.Index - match.Index;
            var relativeLength = innerGroup.Length;
            var cidValue = $"cid:{inlineImage.ContentId}";

            return string.Concat(
                match.Value.AsSpan(0, relativeStart),
                cidValue,
                match.Value.AsSpan(relativeStart + relativeLength));
        });

        return (rewrittenHtml, imagesByPath.Values.ToList());
    }

    private static bool ShouldIgnoreSource(string src)
    {
        if (string.IsNullOrWhiteSpace(src))
        {
            return true;
        }

        return src.StartsWith("cid:", StringComparison.OrdinalIgnoreCase)
            || src.StartsWith("data:", StringComparison.OrdinalIgnoreCase)
            || src.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            || src.StartsWith("https://", StringComparison.OrdinalIgnoreCase);
    }

    private static string? ResolveLocalPath(string src, string baseDirectory)
    {
        var trimmed = src.Trim();
        if (trimmed.Length == 0)
        {
            return null;
        }

        if (Uri.TryCreate(trimmed, UriKind.Absolute, out var absoluteUri))
        {
            if (absoluteUri.IsFile)
            {
                return Path.GetFullPath(Uri.UnescapeDataString(absoluteUri.LocalPath));
            }

            return null;
        }

        var unescaped = Uri.UnescapeDataString(trimmed);
        if (Path.IsPathRooted(unescaped))
        {
            return Path.GetFullPath(unescaped);
        }

        var combined = Path.Combine(baseDirectory, unescaped.Replace('/', Path.DirectorySeparatorChar));
        return Path.GetFullPath(combined);
    }
}
