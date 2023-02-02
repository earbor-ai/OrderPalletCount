
using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OrderPalletCount;
using System.Data;

class Program
{
    static void Main(string[] args)
    {
        string filePath = "C:/Users/pc/Downloads/Earbor_Timesheets/Data.xlsx";
        GetExcelDataTable(filePath);
        Console.ReadLine();
    }
    public static async Task GetExcelDataTable(string filePath)
    {
        DataTable dt = new DataTable();
        List<PalletCalculation> PalletCalculation = new List<PalletCalculation>();
        using (XLWorkbook workBook = new XLWorkbook(filePath))
        {
            IXLWorksheet workSheet = workBook.Worksheet(1);
            bool firstRow = true;
            foreach (IXLRow row in workSheet.Rows())
            {
                if (firstRow)
                {
                    foreach (IXLCell cell in row.Cells())
                    {
                        dt.Columns.Add(cell.Value.ToString());
                    }
                    firstRow = false;
                }
                else
                {
                    dt.Rows.Add();
                    int i = 0;
                    foreach (IXLCell cell in row.Cells())
                    {
                        dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                        i++;
                    }
                    //var data = ConvertToCdrInputs(dt.Rows[i], dt.Columns);
                }
            }
        }
        foreach (DataRow dataRow in dt.Rows)
        {
            PalletCalculation.Add(ConvertToPackageRequestInputs_PalletCalc(dataRow, dt.Columns));
        }

        double palletLenth = 92160;
        PalletCalculation = PalletCalculation.Where(x => x.EventID == "2023000266").ToList();


        var totalInches = PalletCalculation.Sum(x => (double.Parse(x.Width) * double.Parse(x.Height) * double.Parse(x.Length)) * (double.Parse(x.OrderQuantity) / double.Parse(x.CartonQty)));
        var totalWeight = PalletCalculation.Sum(x => double.Parse(x.CartonWeight) * (double.Parse(x.OrderQuantity) / double.Parse(x.CartonQty)));
        var palletsRequired = System.Math.Ceiling(totalInches / palletLenth);

        Console.WriteLine("Total Inches - "+ totalInches);
        Console.WriteLine("Total Weight - "+ totalWeight);
        Console.WriteLine("Pallets Required - "+ palletsRequired);
        Console.ReadLine();
    }
    public static PalletCalculation ConvertToPackageRequestInputs_PalletCalc(DataRow csvInputs, DataColumnCollection headers)
    {
        //ExcelData excelData = new ExcelData();
        var json = @"{}";
        var jObject = JObject.Parse(json);
        for (int i = 0; i < headers.Count; i++)
        {
            jObject[headers[i].ColumnName.Replace(" ", "")] = csvInputs.ItemArray[i].ToString();
        }
        PalletCalculation excelData = JsonConvert.DeserializeObject<PalletCalculation>(jObject.ToString());
        return excelData; //dt.Columns[0].ColumnName
    }
}
