using Newtonsoft.Json;
using System;

namespace WAFSummaryApps.Models
{
    public class BlobListItem
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string Path { get; set; }
        public string LastModified { get; set; }
        public long Size { get; set; }
        public string MediaType { get; set; }
        public bool IsFolder { get; set; }
        public string ETag { get; set; }
        public string FileLocator { get; set; }
        public string LastModifiedBy { get; set; }
    }
}
