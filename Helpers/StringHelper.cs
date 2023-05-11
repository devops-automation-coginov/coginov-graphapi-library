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
    }
}