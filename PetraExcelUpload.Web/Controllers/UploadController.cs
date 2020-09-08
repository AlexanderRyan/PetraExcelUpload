using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace PetraExcelUpload.Web.Controllers
{
    public class UploadController : Controller
    {
        private readonly string[] permittedExtensions = {".xls", ".xlsx"};
        private readonly string targetFilepath;
        private readonly string convertedFilepath;
        private static readonly byte[] XML = { 60, 63, 120, 109, 108, 32 };
        private readonly ILogger<UploadController> logger;
        private int updatedRows;

        public UploadController(IConfiguration config, ILogger<UploadController> logger)
        {
            targetFilepath = config.GetValue<string>("StoredFilesPath");
            convertedFilepath = config.GetValue<string>("ConvertedFilesPath");
            this.logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public async Task<IActionResult> OnPostAsync(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                TempData["Result"] = "No file was selected";

                return RedirectToAction("Index");
            }

            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();

            if (string.IsNullOrEmpty(ext) || !permittedExtensions.Contains(ext))
            {
                TempData["Result"] = $"This file extensions, ({ext}), is not permitted.";

                return RedirectToAction("Index");
            }

            var filePath = Path.Combine(targetFilepath, file.FileName);

            // Get the file signature
            byte[] fileBytes;
            using (var ms = new MemoryStream())
            {
                await file.CopyToAsync(ms);
                fileBytes = ms.ToArray();
                ms.Close();
            }

            System.IO.Directory.CreateDirectory(targetFilepath);

            using (Stream stream = System.IO.File.Create(filePath))
            {
                await file.CopyToAsync(stream);
            }

            // Check if file is raw XML by comparing file signature to XML file sig.
            if (fileBytes.Take(6).SequenceEqual(XML))
            {
                var app = new Microsoft.Office.Interop.Excel.Application();
                var wb = app.Workbooks.Open(Path.GetFullPath(filePath));
                string convertedFilePath = "";

                try
                {
                    wb.SaveAs(Filename: file.FileName + "x", 
                              FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);

                    convertedFilePath = Path.Combine(wb.Path, file.FileName + "x");
                }
                catch (Exception e)
                {
                    logger.LogError(e.ToString());
                }
                finally
                {
                    wb.Close();
                    app.Quit();
                }

                EditExcel(convertedFilePath);
            }
            else
                EditExcel(filePath);

            TempData["Result"] = $"File successfully edited. {updatedRows} rows were edited.";
            TempData["Location"] = $"{Path.GetFullPath(filePath)}";

            return RedirectToAction("Index");
        }
                
        private void EditExcel(string filePath)
        {
            IWorkbook workbook;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
            {
                var ext = Path.GetExtension(filePath);

                if (ext == ".xlsx")
                    workbook = new XSSFWorkbook(fs);
                else
                    workbook = new HSSFWorkbook(fs);

                fs.Close();
            }
            
            ISheet sheet = workbook.GetSheetAt(0);
            int rowCount = sheet.LastRowNum; //? Not being used currently?

            int hourColIndex = 0; //! Stores the index of the column with "Timmar"
            bool colFound = false;

            for (int i = 0; i < sheet.LastRowNum && !colFound; i++)
            {
                IRow row = sheet.GetRow(i);

                if (row == null) continue;

                for (int k = 0; k < row.LastCellNum; k++)
                {
                    ICell cell = row.GetCell(k);
                    if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                    if (cell.ToString().ToLower() == "timmar")
                    {
                        hourColIndex = k;
                        colFound = true;
                        break;
                    }
                }
            }

            for (int i = 4; i < sheet.LastRowNum - 3; i++)
            {
                IRow row = sheet.GetRow(i);
                ICell cell = row.GetCell(row.LastCellNum - 2);

                if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;

                if (cell.ToString().Contains(';'))
                {
                    //? Maybe change code to split the cells text into array here incase more than one semi-colon occurs.
                    //? Create one new row per length of the array to enable multiple new entries instead of only 2.
                    //todo for-loop i< splitArrayn.length, copy newRow to row+1

                    row.CopyRowTo(row.RowNum + 1);
                    IRow newRow = sheet.GetRow(row.RowNum + 1);
                    updatedRows++;

                    for (int j = cell.ColumnIndex; j < row.LastCellNum; j++)
                    {
                        var splitValue = row.GetCell(j).ToString().Split(';');

                        row.GetCell(j).SetCellValue(splitValue[0]);
                        newRow.GetCell(j).SetCellValue(splitValue[1].TrimStart());

                        //! Get the estimated hours for each activity and update the Hour column for each changed row.
                        double rowUpdatedHour = Convert.ToDouble(splitValue[0].Remove(0, splitValue[0].Length - 6).Replace("(","").Replace(")",""));
                        double newRowUpdatedHour = Convert.ToDouble(splitValue[1].Remove(0, splitValue[1].Length - 6).Replace("(", "").Replace(")", ""));

                        row.GetCell(hourColIndex).SetCellValue(rowUpdatedHour);
                        newRow.GetCell(hourColIndex).SetCellValue(newRowUpdatedHour);
                    }
                }
            }

            System.IO.File.Delete(filePath);
            using (FileStream fs = new FileStream(filePath, FileMode.CreateNew, FileAccess.Write))
            {
                workbook.Write(fs);
                fs.Close();
            }
        }
    }
}
