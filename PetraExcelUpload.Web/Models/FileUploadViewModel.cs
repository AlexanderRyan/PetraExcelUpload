using Microsoft.AspNetCore.Http;
using System.ComponentModel.DataAnnotations;

namespace PetraExcelUpload.Web.Models
{
    public class FileUploadViewModel
    {
        [Required]
        public IFormFile File { get; set; }
        public string Result { get; set; }

    }
}
