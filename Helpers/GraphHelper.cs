using Coginov.GraphApi.Library.Models;
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

        public static string GetDriveItemId(this ListItem item)
        {
            return item.DriveItem.Id.Trim();// Fields.AdditionalData.ToDictionary(x => x.Key, x => x.Value);
        }

        public static  string GetDriveId(this ListItem item)
        {
            return item.DriveItem.ParentReference.DriveId;// Fields.AdditionalData.ToDictionary(x => x.Key, x => x.Value);
        }

        public static DriveItemInfo GetDriveItemInfo(this ListItem item)
        {
            return new DriveItemInfo
            {
                DriveId = item.GetDriveId(),
                DriveItemId = item.GetDriveItemId()
            };
        }
    }
}
