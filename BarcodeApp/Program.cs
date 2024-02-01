using NetBarcode;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Drawing;

string fileName = "SampleExcelFile.xlsx";
string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName); // путь к файлу в папке проекта

Console.WriteLine("> BarcodeApp: генерация штрихкода");
try
{

    // Создаем новый файл Excel
    FileInfo newFile = new FileInfo(filePath);
    if (newFile.Exists)
    {
        newFile.Delete();
        newFile = new FileInfo(filePath);
    }

    using (ExcelPackage package = new ExcelPackage(newFile))
    {
        package.DoAdjustDrawings = false;

        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

        Bitmap barcodeImage = GenerateBarcodeImage("EUR000000082");

        //ExcelPicture barcodePicture = worksheet.Drawings.AddPicture("Barcode", barcodeImage);
        //barcodePicture.SetPosition(4, 0, 4, 0);


        package.Save();
    }

    Bitmap GenerateBarcodeImage(string data)
    {
        Barcode barcode = new Barcode();

        Bitmap image = barcode.GetImage(data);

        return image;
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}