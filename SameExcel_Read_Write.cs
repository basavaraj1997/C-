
public class SameExcel_Read_Write{  
public static async Task Main(string[] args)
  {
      Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
      string filepath=Console.ReadLine();
      string strFileName = @""+ filepath;

      Spreadsheet document = new Spreadsheet();
      document.LoadFromFile(strFileName);
      Worksheet worksheet = document.Workbook.Worksheets.ByName("Sheet-1");

      for (int i = 1; i < 1000; i++)
      {
          Cell currentCell = worksheet.Cell(i, 1);
          var query = Convert.ToString(currentCell.Value);
          if (!string.IsNullOrEmpty(query))
          {                    
              string answers = "New Data riting in 'E' columns";
              worksheet.Cell("E" + (i + 1)).Value = answers;
          }
          Console.WriteLine(i);
      }
      document.SaveAs(@"OutputResult.xls");
      document.Close();
      Console.WriteLine("Done");
      Console.ReadKey();
  }
}
