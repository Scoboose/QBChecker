using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace FileComparisonApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string sourceFilePath = string.Empty; // Path to the source CSV file
            string targetFilePath = string.Empty; // Path to the target file
            string resultFilePath = "My_QB_List.xlsx"; // Path to the result Excel file

            if (args.Length < 2)
            {
                Console.WriteLine("Please specify the source file path and target file path as command-line arguments.");
                Console.WriteLine("Example: .\\QBCheck.exe .\\MasterQBList.csv .\\qb.txt");
                Console.ReadLine();
                return;
            }

            sourceFilePath = args[0];
            targetFilePath = args[1];

            List<SourceData> sourceData = ReadSourceFile(sourceFilePath);
            List<string> targetNames = ReadTargetFile(targetFilePath);

            List<ComparisonResult> comparisonResults = CompareData(sourceData, targetNames);
            List<string> unknownNames = GetUnknownNames(targetNames, comparisonResults);

            GenerateExcelFile(resultFilePath, sourceData, comparisonResults, unknownNames);

            //WriteColor("This is my message with new color with red", ("{message}", ConsoleColor.Red), ("{with}", ConsoleColor.Blue));
            Console.WriteLine("");
            WriteColor("Comparison completed! My_QB_List.xlsx generated.", ("{completed!}", ConsoleColor.Green));
            Console.WriteLine("");
            Console.WriteLine("");
            WriteColor("Created by Pokey for use on the Infinite Leafetide Asheron's call server", ("{Pokey}", ConsoleColor.Blue), ("{Infinite}", ConsoleColor.Red), ("{Leafetide}", ConsoleColor.Red));
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("You may now close this window.");
            Console.ReadLine();
        }

        static List<SourceData> ReadSourceFile(string filePath)
        {
            List<SourceData> sourceData = new List<SourceData>();

            using (StreamReader reader = new StreamReader(filePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    string[] parts = line.Split(',');
                    if (parts.Length >= 5)
                    {
                        string name = parts[0].Trim();
                        string quest = parts[1].Trim();
                        string action = parts[2].Trim();
                        string notes = parts[3].Trim();
                        string serverSideDefinition = parts[4].Trim();
                        sourceData.Add(new SourceData(name, quest, action, notes, serverSideDefinition));
                    }
                }
            }

            return sourceData;
        }

        static List<string> ReadTargetFile(string filePath)
        {
            return File.ReadAllLines(filePath).Select(line => line.Trim()).ToList();
        }

        static List<ComparisonResult> CompareData(List<SourceData> sourceData, List<string> targetNames)
        {
            List<ComparisonResult> comparisonResults = new List<ComparisonResult>();

            foreach (SourceData data in sourceData)
            {
                bool isMatch = targetNames.Contains(data.Name);
                ComparisonResult result = new ComparisonResult(data.Name, data.Quest, data.Action, data.Notes, data.ServerSideDefinition, isMatch);
                comparisonResults.Add(result);
            }

            return comparisonResults;
        }

        static List<string> GetUnknownNames(List<string> targetNames, List<ComparisonResult> comparisonResults)
        {
            List<string> unknownNames = targetNames.Except(comparisonResults.Select(c => c.Name)).ToList();
            return unknownNames;
        }

        static void GenerateExcelFile(string filePath, List<SourceData> sourceData, List<ComparisonResult> results, List<string> unknownNames)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Comparison Results");

                // Set column headers
                worksheet.Cells[1, 1].Value = "Stamp Name";
                worksheet.Cells[1, 2].Value = "Quest";
                worksheet.Cells[1, 3].Value = "Action";
                worksheet.Cells[1, 4].Value = "Notes";
                worksheet.Cells[1, 5].Value = "Server Side Definition";
                worksheet.Cells[1, 6].Value = "Compleated";

                // Populate data
                for (int i = 0; i < results.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = sourceData[i].Name;
                    worksheet.Cells[i + 2, 2].Value = sourceData[i].Quest;
                    worksheet.Cells[i + 2, 3].Value = sourceData[i].Action;
                    worksheet.Cells[i + 2, 4].Value = sourceData[i].Notes;
                    worksheet.Cells[i + 2, 5].Value = sourceData[i].ServerSideDefinition;
                    worksheet.Cells[i + 2, 6].Value = results[i].IsMatch;
                }

                // Create sheet for unknown names
                if (unknownNames.Count > 0)
                {
                    ExcelWorksheet unknownWorksheet = package.Workbook.Worksheets.Add("Unknown Stamps");

                    // Set column headers
                    unknownWorksheet.Cells[1, 1].Value = "Stamps you have that are not on the master QB list";

                    // Populate data
                    for (int i = 0; i < unknownNames.Count; i++)
                    {
                        unknownWorksheet.Cells[i + 2, 1].Value = unknownNames[i];
                    }
                }

                // Save the Excel file
                FileInfo file = new FileInfo(filePath);
                package.SaveAs(file);
            }
        }

        private static void WriteColor(string str, params (string substring, ConsoleColor color)[] colors)
        {
            var words = Regex.Split(str, @"( )");

            foreach (var word in words)
            {
                (string substring, ConsoleColor color) cl = colors.FirstOrDefault(x => x.substring.Equals("{" + word + "}"));
                if (cl.substring != null)
                {
                    Console.ForegroundColor = cl.color;
                    Console.Write(cl.substring.Substring(1, cl.substring.Length - 2));
                    Console.ResetColor();
                }
                else
                {
                    Console.Write(word);
                }
            }
        }
    }

    class SourceData
    {
        public string Name { get; set; }
        public string Quest { get; set; }
        public string Action { get; set; }
        public string Notes { get; set; }
        public string ServerSideDefinition { get; set; }

        public SourceData(string name, string quest, string action, string notes, string serverSideDefinition)
        {
            Name = name;
            Quest = quest;
            Action = action;
            Notes = notes;
            ServerSideDefinition = serverSideDefinition;
        }
    }

    class ComparisonResult
    {
        public string Name { get; set; }
        public string Quest { get; set; }
        public string Action { get; set; }
        public string Notes { get; set; }
        public string ServerSideDefinition { get; set; }
        public bool IsMatch { get; set; }

        public ComparisonResult(string name, string quest, string action, string notes, string serverSideDefinition, bool isMatch)
        {
            Name = name;
            Quest = quest;
            Action = action;
            Notes = notes;
            ServerSideDefinition = serverSideDefinition;
            IsMatch = isMatch;
        }
    }
}