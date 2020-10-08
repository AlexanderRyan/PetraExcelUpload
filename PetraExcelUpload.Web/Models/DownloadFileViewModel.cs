using System.Collections.Generic;

namespace PetraExcelUpload.Web.Models
{
    public class DownloadFileViewModel
    {
        public IList<BlobFileViewModel> BlobFiles { get; set; } = new List<BlobFileViewModel>();
    }
}
