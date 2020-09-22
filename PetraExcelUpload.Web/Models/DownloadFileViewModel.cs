using System.Collections.Generic;

namespace PetraExcelUpload.Web.Models
{
    public class DownloadFileViewModel
    {
        //public string FileUrl { get; set; }

        //public IList<BlobItem> BlobFiles { get; set; } = new List<BlobItem>();

        public IList<BlobFileViewModel> BlobFiles { get; set; } = new List<BlobFileViewModel>();
    }
}
