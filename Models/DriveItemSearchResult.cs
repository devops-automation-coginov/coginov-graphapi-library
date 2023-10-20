using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;

namespace Coginov.GraphApi.Library.Models
{
    public class DriveItemSearchResult
    {
        public List<DriveItem> DocumentIds { get; set; }
        public string SkipToken { get; set; }
        public DateTime LastDate {get; set; }
        public bool HasMoreResults { get; set; }
    }
}