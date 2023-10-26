using Microsoft.Graph.Models;
using System.Collections.Generic;
using System.Linq;

namespace Coginov.GraphApi.Library.Helpers
{
    public static class GraphHelper
    {
        public static Dictionary<string, object> GetFieldValues(this ListItem item)
        {
            return item.Fields.AdditionalData.ToDictionary(x => x.Key, x => x.Value);
        }
    }
}
