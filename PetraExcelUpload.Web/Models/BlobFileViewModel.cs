using Azure.Storage.Blobs.Models;

namespace PetraExcelUpload.Web.Models
{
    public class BlobFileViewModel
    {
        public string Url { get; set; }
        public BlobItem File { get; set; }
    }
}
