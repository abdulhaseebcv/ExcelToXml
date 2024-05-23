using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Xml;

public class HomeController : Controller
{
    [HttpGet]
    public IActionResult Index(string message = null, string xmlFileName = null)
    {
        var model = (message, xmlFileName);
        return View(model);
    }

    [HttpPost]
    public IActionResult Upload([FromForm] Microsoft.AspNetCore.Http.IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return RedirectToAction("Index", new { message = "No file uploaded." });
        }

        if (!(file.FileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
              file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)))
        {
            return RedirectToAction("Index", new { message = "Invalid file format. Only XLS and XLSX files are supported." });
        }

        try
        {
            var result = ProcessFileAndGenerateXml(file);
            return RedirectToAction("Index", new { message = result.message, xmlFileName = result.xmlFileName });
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"An error occurred while processing the file: {ex.Message}");
        }
    }

    public IActionResult DownloadXml(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
            return NotFound();

        string xmlFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", fileName);

        if (!System.IO.File.Exists(xmlFilePath))
            return NotFound();

        var fileStream = new FileStream(xmlFilePath, FileMode.Open);
        return File(fileStream, "application/xml", fileName);
    }

    private (string message, string xmlFileName) ProcessFileAndGenerateXml(Microsoft.AspNetCore.Http.IFormFile file)
    {
        string xmlFileName = $"output_{DateTime.Now:yyyyMMddHHmmssfff}.xml";
        string xmlFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", xmlFileName);

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

            xmlDoc.Save(xmlFilePath);
        }

        return ("File uploaded and processed successfully.", xmlFileName);
    }
}
