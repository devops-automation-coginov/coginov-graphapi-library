using System;
using System.IO;

namespace Coginov.GraphApi.Library.Helpers
{
    public static class StringHelper
    {
        public static string GetFolderNameFromSpoUrl(this string url)
        {
            var uri = new Uri(url);
            var subsite = uri.PathAndQuery.TrimStart('/').Replace("sites/", "");
            return $"{uri.Host}-{subsite}";
        }

        public static string GetFilePathWithTimestamp(this string filePath)
        {
            try
            {
                var timeStamp = $".{DateTime.Now.ToString("yyyyMMddHHmmssfff")}";

                var fDir = Path.GetDirectoryName(filePath);
                var fName = Path.GetFileNameWithoutExtension(filePath);
                var fExt = Path.GetExtension(filePath);

                return Path.Combine(fDir, String.Concat(fName, timeStamp, fExt));
            }
            catch(Exception)
            {
                return filePath;
            }
        }

        public static bool IsRootUrl(this string url)
        {
            var uri = new Uri(url);
            var path = uri.GetComponents(UriComponents.Path, UriFormat.Unescaped);
            return path.Length == 0;
        }

        public static string ExtractStringAfterRoot(this string input)
        {
            // Find the index of "root:"
            int rootIndex = input.ToLower().IndexOf("root:", StringComparison.OrdinalIgnoreCase);

            if (rootIndex >= 0)
            {
                // Extract the substring after "root:"
                return input.Substring(rootIndex + 5); // 5 is the length of "root:"
            }

            // "root:" not found in the input
            return string.Empty; // Return an empty string
        }
    }
}