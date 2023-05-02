namespace Coginov.GraphApi.Library.Models
{
    public class DriveConnectionInfo
    {
        public string Id { get; set; }
        public string Root { get; set; }
        public string Path { get; set; }
        public string Name { get; set; }
        public string GroupId { get; set; } 
        public bool DownloadCompleted {get; set;}
    }
}