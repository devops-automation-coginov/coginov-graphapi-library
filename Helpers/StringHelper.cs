using System;
using System.IO;
using System.Linq;

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
    }
}