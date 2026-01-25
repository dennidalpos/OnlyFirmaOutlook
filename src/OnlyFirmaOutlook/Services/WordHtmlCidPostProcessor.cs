using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using OnlyFirmaOutlook.Models;

namespace OnlyFirmaOutlook.Services;

public static class WordHtmlCidPostProcessor
{
    private static readonly Regex ImageSrcRegex = new(
        @"<\s*(?<tag>img|v:imagedata)\b[^>]*?\bsrc\s*=\s*(?<value>('(?<inner>[^']*)'|""(?<inner>[^""]*)""|(?<inner>[^\s>]+)))",
        RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled);
    private static readonly Regex CssUrlRegex = new(
        @"url\(\s*(?<value>('(?<inner>[^']*)'|""(?<inner>[^""]*)""|(?<inner>[^)\s]+)))\s*\)",
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
        if (File.Exists(normalizedBaseDirectory))
        {
            normalizedBaseDirectory = Path.GetDirectoryName(normalizedBaseDirectory)
                ?? normalizedBaseDirectory;
        }
        var imagesByPath = new Dictionary<string, InlineImage>(StringComparer.OrdinalIgnoreCase);

        string rewrittenHtml = ReplaceLocalSourcesWithCid(
            html,
            normalizedBaseDirectory,
            imagesByPath,
            ImageSrcRegex,
            isCssUrl: false);
        rewrittenHtml = ReplaceLocalSourcesWithCid(
            rewrittenHtml,
            normalizedBaseDirectory,
            imagesByPath,
            CssUrlRegex,
            isCssUrl: true);

        return (rewrittenHtml, imagesByPath.Values.ToList());
    }

    private static string ReplaceLocalSourcesWithCid(
        string html,
        string baseDirectory,
        Dictionary<string, InlineImage> imagesByPath,
        Regex regex,
        bool isCssUrl)
    {
        return regex.Replace(html, match =>
        {
            var srcValue = match.Groups["inner"].Value;
            if (isCssUrl)
            {
                srcValue = NormalizeCssUrl(srcValue);
            }
            if (ShouldIgnoreSource(srcValue))
            {
                return match.Value;
            }

            var resolvedPath = ResolveLocalPath(srcValue, baseDirectory);
            if (resolvedPath is null)
            {
                return match.Value;
            }

            if (!imagesByPath.TryGetValue(resolvedPath, out var inlineImage))
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
    }

    private static string NormalizeCssUrl(string value)
    {
        return Regex.Replace(value, @"\\(.)", "$1");
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

        var decoded = WebUtility.HtmlDecode(trimmed);
        if (TryResolveFileUri(decoded, out var fileUriPath))
        {
            return fileUriPath;
        }

        var candidates = new List<string>();
        AddAbsoluteCandidates(decoded, candidates);
        AddRelativeCandidates(decoded, baseDirectory, candidates);

        if (!string.Equals(decoded, trimmed, StringComparison.Ordinal))
        {
            AddAbsoluteCandidates(trimmed, candidates);
            AddRelativeCandidates(trimmed, baseDirectory, candidates);
        }

        foreach (var candidate in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }

    private static bool TryResolveFileUri(string value, out string? path)
    {
        path = null;
        if (!value.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var withoutScheme = value[5..];
        if (withoutScheme.StartsWith("///", StringComparison.Ordinal))
        {
            withoutScheme = withoutScheme[3..];
        }
        else if (withoutScheme.StartsWith("//", StringComparison.Ordinal))
        {
            withoutScheme = withoutScheme[2..];
        }

        withoutScheme = withoutScheme.TrimStart('/');
        var filePath = Uri.UnescapeDataString(withoutScheme)
            .Replace('/', Path.DirectorySeparatorChar)
            .Replace('\\', Path.DirectorySeparatorChar);

        if (!Path.IsPathRooted(filePath))
        {
            return false;
        }

        var fullPath = Path.GetFullPath(filePath);
        if (File.Exists(fullPath))
        {
            path = fullPath;
            return true;
        }

        return false;
    }

    private static void AddAbsoluteCandidates(string value, ICollection<string> candidates)
    {
        if (Uri.TryCreate(value, UriKind.Absolute, out var absoluteUri) && absoluteUri.IsFile)
        {
            candidates.Add(Path.GetFullPath(Uri.UnescapeDataString(absoluteUri.LocalPath)));
        }

        var unescaped = Uri.UnescapeDataString(value);
        if (Path.IsPathRooted(unescaped))
        {
            candidates.Add(Path.GetFullPath(unescaped));
        }
    }

    private static void AddRelativeCandidates(string value, string baseDirectory, ICollection<string> candidates)
    {
        var unescaped = Uri.UnescapeDataString(value);
        var combined = Path.Combine(baseDirectory, unescaped.Replace('/', Path.DirectorySeparatorChar));
        candidates.Add(Path.GetFullPath(combined));
    }
}
