using Newtonsoft.Json;

namespace Coginov.GraphApi.Library.Models
{
    public class RecycleBinItem
    {
        [JsonProperty("id")]
        public string? Id { get; set; }

        [JsonProperty("name")]
        public string? Name { get; set; }

        [JsonProperty("driveItemId")]
        public string? DriveItemId { get; set; }
    }
}
