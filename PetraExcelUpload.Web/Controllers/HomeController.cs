using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Azure.Storage.Blobs;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using PetraExcelUpload.Web.Models;

namespace PetraExcelUpload.Web.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment webHostEnvironment;
        private readonly ILogger<HomeController> logger;

        private readonly string[] permittedExtensions = {".xls", ".xlsx"};
        private static readonly byte[] XML = { 60, 63, 120, 109, 108, 32 };

        [BindProperty]
        public FileUploadViewModel ViewModel { get; set; }
        public IConfiguration Config { get; }

        public HomeController(ILogger<HomeController> logger, 
                                IWebHostEnvironment webHostEnvironment,
                                IConfiguration config)
        {
            this.logger = logger;
            this.webHostEnvironment = webHostEnvironment;
            Config = config;
        }

        public IActionResult Index()
        {
            ViewModel = new FileUploadViewModel();

            return View(ViewModel);
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid)
            {
                ViewModel.Result = "Please choose a file.";

                return View("Index", ViewModel);
            }

            if (ViewModel.File == null || ViewModel.File.Length == 0)
            {
                ViewModel.Result = "No file was selected";

                return View("Index", ViewModel);
            }

            var ext = Path.GetExtension(ViewModel.File.FileName).ToLowerInvariant();

            if (string.IsNullOrEmpty(ext) || !permittedExtensions.Contains(ext))
            {
                ViewModel.Result = $"This file extensions, ({ext}), is not permitted.";

                return View("Index", ViewModel);
            }

            string standardFilePath = Path.Combine(webHostEnvironment.WebRootPath, "files");
            string convFilePath = Path.Combine(webHostEnvironment.WebRootPath, "temp"); //! Path for file to be converted
            string filePath;

            byte[] fileBytes;
            using (var ms = new MemoryStream())
            {
                await ViewModel.File.CopyToAsync(ms);
                fileBytes = ms.ToArray();
                ms.Close();
            }

            //! If File-sequence matches XML, file is old format and needs to be 'saved as' new format with Office Interop Excel.
            if (fileBytes.Take(6).SequenceEqual(XML))
            {
                System.IO.Directory.CreateDirectory(convFilePath);
                filePath = Path.Combine(convFilePath, ViewModel.File.FileName);

                using (Stream stream = System.IO.File.Create(filePath))
                {
                    await ViewModel.File.CopyToAsync(stream);
                }

                //! Open new Excel app and open the file with old format in a new workbook.
                var app = new Microsoft.Office.Interop.Excel.Application();
                var wb = app.Workbooks.Open(filePath);

                try
                { 
                    //! Save the old format Excel to wwwroot/files as new excel with new format - .xlsx
                    string fileName = Path.Combine(standardFilePath, ViewModel.File.FileName + "x");
                    wb.SaveAs(Filename: fileName,
                    FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                    
                    System.IO.File.Delete(filePath); //! Delete the temp file stored for conversion.
                    filePath = fileName;
                }
                catch (Exception error)
                {
                    logger.LogError(error.ToString());
                }
                finally
                {
                    ForceExitExcel();
                    wb.Close();
                    app.Quit();
                }

                UploadToBlobStorage("petra-excel", ViewModel.File.FileName + "x", filePath);
                ForceExitExcel();
            }
            else
            {
                System.IO.Directory.CreateDirectory(standardFilePath);
                filePath = Path.Combine(standardFilePath, ViewModel.File.FileName);

                using (Stream stream = System.IO.File.Create(filePath))
                {
                    await ViewModel.File.CopyToAsync(stream);
                }

                UploadToBlobStorage("petra-excel", ViewModel.File.FileName, filePath);
            }
            System.Threading.Thread.Sleep(8000); //? Ugly solution - Azure function doesnt finish before page reloads resulting in Downloads displaying no files uploaded.
            return RedirectToAction("Index", "Download");
        }

        private void UploadToBlobStorage(string cntName, string fileName, string filePath)
        {
            string accessKey = Config.GetConnectionString("AccessKey");

            BlobContainerClient container = new BlobContainerClient(accessKey, cntName);

            container.CreateIfNotExists();

            BlobClient blobClient = container.GetBlobClient("uploads/" + fileName);
            
            using FileStream uploadFileStream = System.IO.File.OpenRead(filePath);
            blobClient.Upload(uploadFileStream, true);
            
            uploadFileStream.Close();

            //! Deletes the file from the 'wwwroot/files' folder since it's already uploaded to blob storage.
            System.IO.File.Delete(filePath);
        }

        private void ForceExitExcel()
        {
            //! Makes sure COM references are released.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
