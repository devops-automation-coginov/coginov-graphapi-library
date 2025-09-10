using Newtonsoft.Json;
using System.Collections.Generic;

namespace Coginov.GraphApi.Library.Models
{
    public class RecycleBinCollectionResponse
    {
        [JsonProperty("value")]
        public List<RecycleBinItem> Value { get; set; } = new();
    }
}
