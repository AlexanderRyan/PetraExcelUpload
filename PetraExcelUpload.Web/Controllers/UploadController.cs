using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace PetraExcelUpload.Web.Controllers
{
    public class UploadController : Controller
    {
        private readonly IWebHostEnvironment webHostEnvironment;
        private readonly ILogger<UploadController> logger;

        private readonly string[] permittedExtensions = {".xls", ".xlsx"};
        private static readonly byte[] XML = { 60, 63, 120, 109, 108, 32 };

        private int nrOfUpdatedRows;

        public UploadController(ILogger<UploadController> logger, IWebHostEnvironment webHostEnvironment)
        {
            this.logger = logger;
            this.webHostEnvironment = webHostEnvironment;
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
            
            byte[] fileBytes;
            
            string standardFilePath = Path.Combine(webHostEnvironment.WebRootPath, "files");
            string convFilePath = Path.Combine(webHostEnvironment.WebRootPath, "temp"); //! Path for file to be converted
            string filePath;

            using (var ms = new MemoryStream())
            {
                await file.CopyToAsync(ms);
                fileBytes = ms.ToArray();
                ms.Close();
            }

            if (fileBytes.Take(6).SequenceEqual(XML))
            {
                System.IO.Directory.CreateDirectory(convFilePath);
                filePath = Path.Combine(convFilePath, file.FileName);

                using (Stream stream = System.IO.File.Create(filePath))
                {
                    await file.CopyToAsync(stream);
                }

                //! Open new Excel app and open the file with old format in a new workbook.
                var app = new Microsoft.Office.Interop.Excel.Application();
                var wb = app.Workbooks.Open(filePath);

                try
                { 
                    //! Save the old format Excel to wwwroot/files as new excel with new format - .xlsx
                    string fileName = Path.Combine(standardFilePath, file.FileName + "x");
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

            }
            else
            {
                System.IO.Directory.CreateDirectory(standardFilePath);

                filePath = Path.Combine(standardFilePath, file.FileName);
                using (Stream stream = System.IO.File.Create(filePath))
                {
                    await file.CopyToAsync(stream);
                }

            }

            EditExcel(filePath);

            TempData["Result"] = $"File successfully edited. {nrOfUpdatedRows} rows were edited.";
            TempData["Location"] = $"{Path.GetFullPath(filePath)}";

            ForceExitExcel();

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
            short rowHeight = sheet.GetRow(4).Height; //! Get the original row height
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

                        if (j == cell.ColumnIndex) //! If cell is "EXTERN NOTERING" - Copy new rows and update nr of new rows.
                        {
                            for (int k = 1; k < splitCellValues.Length; k++)
                            {
                                row.CopyRowTo(row.RowNum + k);
                                nrOfUpdatedRows++;
                                i++; //! No need to check newly added rows, should already be formatted correctly.
                            }

                            //! 'EXTERN & INTERN Notering' always ends with specified hours formattade as (X.XX) which equals 6 characters.
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

            //! Reformat all the rows back to original height, CopyRowTo somehow breaks the rowHeight of the succeeding row;
            for (int i = 1; i < sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);

                if (row == null) continue;

                row.Height = rowHeight;
            }

            System.IO.File.Delete(filePath);
            using (FileStream fs = new FileStream(filePath, FileMode.CreateNew, FileAccess.Write))
            {
                workbook.Write(fs);
                fs.Close();
            }
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
