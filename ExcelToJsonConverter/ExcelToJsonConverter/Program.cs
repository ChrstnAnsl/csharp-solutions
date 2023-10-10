using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        // Replace with your Excel file path
        string excelFilePath = "test.xlsx";

        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
        {
            WorkbookPart? workbookPart = spreadsheetDocument.WorkbookPart;

            if (workbookPart != null)
            {
                WorksheetPart? worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();

                if (worksheetPart != null)
                {
                    // Load the textual content from the worksheet
                    string excelContent = ReadExcelContent(worksheetPart, workbookPart);

                    Console.WriteLine(excelContent);

                    // Generate JSON content
                    string jsonContent = GenerateJsonContent(excelContent);

                    // Save JSON content to a file
                    string jsonFilePath = "output.json"; // Replace with your desired output file path
                    File.WriteAllText(jsonFilePath, jsonContent);

                    //Console.WriteLine(jsonContent);

                    // Console.WriteLine("JSON content saved to " + jsonFilePath);
                }
                else
                {
                    Console.WriteLine("No worksheet found in the Excel file.");
                }
            }
            else
            {
                Console.WriteLine("No workbook found in the Excel file.");
            }
        }
    }

    static string ReadExcelContent(WorksheetPart worksheetPart, WorkbookPart workbookPart)
    {
        string excelContent = "";

        if (worksheetPart != null)
        {
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            foreach (Row row in sheetData.Elements<Row>())
            {
                foreach (Cell cell in row.Elements<Cell>())
                {
                    string cellValue = GetCellValue(workbookPart, cell);
                    excelContent += cellValue + ";"; // Separate cell values with tabs
                }
                excelContent += Environment.NewLine; // Start a new line for each row
            }
        }

        return excelContent;
    }

    static string GetCellValue(WorkbookPart workbookPart, Cell cell)
    {
        if (cell == null || cell.CellValue == null)
        {
            return "";
        }

        string cellValue = cell.CellValue.Text;

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            // If the cell contains a shared string, look up the shared string value
            SharedStringTablePart? sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (sharedStringTablePart != null)
            {
                int index;
                if (int.TryParse(cellValue, out index))
                {
                    SharedStringItem sharedStringItem = sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(index);

                    if (sharedStringItem?.Text?.Text != null)
                    {
                        cellValue = sharedStringItem.Text.Text;
                    }
                    
                }
            }
        }

        return cellValue;
    }

    static string GenerateJsonContent(string excelContent)
    {
        var jsonResult = new List<TestCase>();
        var lines = excelContent.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
        var currentTestCase = new TestCase();
        var inSteps = false;

        foreach (var line in lines)
        {
            if (line.StartsWith("Test Case") || line.StartsWith("New Functionality") || line.StartsWith("Regression") || line.StartsWith("StandardChecklist"))
            {
                // If we encounter a new test case, add the previous one (if not empty) and start a new one
                if (!string.IsNullOrEmpty(currentTestCase.Name))
                {
                    // Filter out empty strings from Steps and Results before adding the test case
                    currentTestCase.Steps = currentTestCase.Steps.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                    currentTestCase.Results = currentTestCase.Results.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                    jsonResult.Add(currentTestCase);
                }

                currentTestCase = new TestCase
                {
                    Name = line.Trim(),
                    Steps = new List<string>(),
                    Results = new List<string>()
                };

                inSteps = false;
            }
            else if (!string.IsNullOrWhiteSpace(line))
            {
                if (line.StartsWith("Step"))
                {
                    inSteps = true;
                }
                else if (line.StartsWith("Result"))
                {
                    inSteps = false;
                }
                else if (inSteps)
                {
                    // Split the line using the semicolon (';') as a delimiter and add to the Steps list
                    var stepParts = line.Split(';').Select(step => step.Trim());
                    currentTestCase.Steps.AddRange(stepParts.Where(s => !string.IsNullOrWhiteSpace(s)));
                }
                else
                {
                    // Split the line using the semicolon (';') as a delimiter and add to the Results list
                    var resultParts = line.Split(';').Select(result => result.Trim());
                    currentTestCase.Results.AddRange(resultParts.Where(s => !string.IsNullOrWhiteSpace(s)));
                }
            }
        }

        // Add the last test case (if not empty)
        if (!string.IsNullOrEmpty(currentTestCase.Name))
        {
            // Filter out empty strings from Steps and Results before adding the test case
            currentTestCase.Steps = currentTestCase.Steps.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
            currentTestCase.Results = currentTestCase.Results.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
            jsonResult.Add(currentTestCase);
        }

        // Serialize to JSON
        string json = JsonConvert.SerializeObject(jsonResult, Formatting.Indented);

        return json;
    }

}
public class TestCase
{
    public string Name { get; set; }
    public List<string> Steps { get; set; }
    public List<string> Results { get; set; }
}