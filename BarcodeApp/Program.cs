using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Drawing;
using System.Net;

string fileName = "SampleExcelFile.xlsx";
string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName); // путь к файлу в папке проекта

string imageUrl = "https://barcodeapi.org/api/128/EUR000000082";

Console.WriteLine("> BarcodeApp: генерация штрихкода")ж

// Создаем новый файл Excel
FileInfo newFile = new FileInfo(filePath);
if (newFile.Exists)
{
    newFile.Delete();
    newFile = new FileInfo(filePath);
}

using (ExcelPackage package = new ExcelPackage(newFile))
{
    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

    using (WebClient client = new WebClient())
    {
        byte[] imageData = client.DownloadData(imageUrl);

        using (MemoryStream memoryStream = new MemoryStream(imageData))
        {
            using (Bitmap bitmap = new Bitmap(memoryStream))
            {
                // Добавление изображения в документ Excel
                ExcelPicture picture = worksheet.Drawings.AddPicture("barcodeImage", bitmap);
                picture.SetPosition(1, 0, 1, 0);
            }
        }
    }

    package.Save();
}