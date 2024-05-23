using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.IO;
using System.Xml;

public class HomeController : Controller
{
    [HttpGet]
    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public IActionResult Upload([FromForm] Microsoft.AspNetCore.Http.IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return View("Index", "No file uploaded.");
        }

        if (!(file.FileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
              file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)))
        {
            return View("Index", "Invalid file format. Only XLS and XLSX files are supported.");
        }

        try
        {
            ProcessFileAndGenerateXml(file);
            return View("Index", "File uploaded and processed successfully.");
        }
        catch (Exception)
        {
            return StatusCode(500, "An error occurred while processing the file.");
        }
    }

    private void ProcessFileAndGenerateXml(Microsoft.AspNetCore.Http.IFormFile file)
    {
        using (var stream = new MemoryStream())
        {
            file.CopyTo(stream);
            stream.Position = 0;

            IWorkbook workbook;
            if (file.FileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
            {
                workbook = new HSSFWorkbook(stream);
            }
            else
            {
                workbook = new XSSFWorkbook(stream);
            }

            ISheet sheet = workbook.GetSheetAt(0);
            var xmlDoc = new XmlDocument();
            xmlDoc.AppendChild(xmlDoc.CreateElement("Data"));

            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                IRow sheetRow = sheet.GetRow(row);
                if (sheetRow != null)
                {
                    var rowNode = xmlDoc.CreateElement("Row");

                    for (int col = 0; col < sheetRow.LastCellNum; col++)
                    {
                        ICell cell = sheetRow.GetCell(col);
                        var cellValue = cell?.ToString() ?? "";
                        var colNode = xmlDoc.CreateElement("Column" + (col + 1));
                        colNode.InnerText = cellValue;
                        rowNode.AppendChild(colNode);
                    }

                    xmlDoc.DocumentElement.AppendChild(rowNode);
                }
            }

            string xmlFileName = $"output_{DateTime.Now:yyyyMMddHHmmssfff}.xml";
            string xmlFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", xmlFileName);
            xmlDoc.Save(xmlFilePath);
        }
    }
}
