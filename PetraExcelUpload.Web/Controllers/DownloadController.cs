using Azure.Storage.Blobs;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using PetraExcelUpload.Web.Models;
using System.Threading.Tasks;

namespace PetraExcelUpload.Web.Controllers
{
    public class DownloadController : Controller
    {
        private readonly ILogger<DownloadController> logger;
        public IConfiguration Config { get; }
        public DownloadController(ILogger<DownloadController> logger,
                                IConfiguration config)
        {
            this.logger = logger;
            Config = config;
        }
        public async Task<IActionResult> Index()
        {
            string accessKey = Config.GetConnectionString("AccessKey");
            BlobContainerClient container = new BlobContainerClient(accessKey, "petra-excel");

            DownloadFileViewModel downloadFileVM = new DownloadFileViewModel();

            await foreach (var item in container.GetBlobsAsync(prefix: "downloads/"))
            {
                BlobFileViewModel blob = new BlobFileViewModel
                {
                    File = item,
                    Url = container.Uri.AbsoluteUri + "/" + item.Name
                };
                downloadFileVM.BlobFiles.Add(blob);
            }

            return View(downloadFileVM);
        }
        public async Task<IActionResult> Delete()
        {
            string accessKey = Config.GetConnectionString("AccessKey");
            BlobContainerClient container = new BlobContainerClient(accessKey, "petra-excel");

            await foreach (var item in container.GetBlobsAsync(prefix: "downloads/"))
                await container.DeleteBlobIfExistsAsync(item.Name);

            logger.LogInformation($"Uploaded blobs removed from Azure Storage Container.");
            return RedirectToAction("Index");
        }
    }
}
