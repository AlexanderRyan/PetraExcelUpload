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
        private readonly ILogger<UploadController> logger;
        private int updatedRows;

        public UploadController(IConfiguration config, ILogger<UploadController> logger)
        {
            targetFilepath = config.GetValue<string>("StoredFilesPath");
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

            System.IO.Directory.CreateDirectory(targetFilepath);

            using (Stream stream = System.IO.File.Create(filePath))
            {
                await file.CopyToAsync(stream);
            }

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
                    for (int j = cell.ColumnIndex; j < row.LastCellNum; j++)
                    {
                        var splitCellValues = row.GetCell(j).ToString().Split(";");

                        if (j == cell.ColumnIndex) //! If cell is "EXTERN NOTERING" - Kopiera nya rader och uppdatera antal nya rader
                        {
                            for (int k = 1; k < splitCellValues.Length; k++)
                            {
                                row.CopyRowTo(row.RowNum + k);
                                updatedRows++;
                                i++; //! No need to check newly added rows, should already be formatted correctly.
                            }

                            double rowUpdatedHour = Convert.ToDouble(splitCellValues[0]
                                .Remove(0, splitCellValues[0].Length - 6).Replace("(", "").Replace(")", ""));

                            row.GetCell(j).SetCellValue(splitCellValues[0]); //! Updates the orignal cell to desired content.
                            row.GetCell(hourColIndex).SetCellValue(rowUpdatedHour); //! Updates original cells Hour-Column

                            //! Loop sets the added rows content and their Hour-Column
                            for (int k = 1; k < splitCellValues.Length; k++)
                            {
                                IRow newRow = sheet.GetRow(row.RowNum + k);
                                rowUpdatedHour = Convert.ToDouble(splitCellValues[k].Remove(0, splitCellValues[k].Length - 6).Replace("(", "").Replace(")", ""));

                                newRow.GetCell(j).SetCellValue(splitCellValues[k].TrimStart());
                                newRow.GetCell(hourColIndex).SetCellValue(rowUpdatedHour);
                            }
                        }
                        else //! If cell is NOT "EXTERN NOTERING", it will be "INTERN NOTERING" - Only set new cell content
                        {
                            row.GetCell(j).SetCellValue(splitCellValues[0]); //! Updates the orignal cell to desired content.

                            for (int k = 1; k < splitCellValues.Length; k++)
                            {
                                IRow newRow = sheet.GetRow(row.RowNum + k);
                                newRow.GetCell(j).SetCellValue(splitCellValues[k].TrimStart());
                            }
                        }
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
