using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;


namespace TryCSV2ODSConvertion
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Process starts.");

            var directory = @"D:\CBS_Work\TeamOpenData\TryCSV2ODSConvertion\Data\";
            var tempDir = Path.Combine(directory, "tempCSV");

            string[] csvFiles = Directory.GetFiles(directory, "*.csv");

            // Constant row count limit
            int MAX_ROW_COUNT = 1048576;

            // Split the large csv file
            foreach (string csvFile in csvFiles)
            {
                using (StreamReader stReader = new StreamReader(csvFile))
                {
                    int fileIndex = 0;

                    while (!stReader.EndOfStream)
                    {
                        int count = 0;
                        using (StreamWriter stWriter = new StreamWriter(Path.Combine(tempDir, fileIndex++ + Path.GetFileName(csvFile))))
                        {
                            stWriter.AutoFlush = true;
                            while (!stReader.EndOfStream && ++count < MAX_ROW_COUNT)
                            {
                                stWriter.WriteLine(stReader.ReadLine());
                            }
                        }
                    }
                }
            }

            // Convert smaller temp csv files
            string[] smallerCSVs = Directory.GetFiles(tempDir, "*.csv");

            string outputODS = Path.Combine(tempDir, "output.ods");

            // load the csv file in an instance of Workbook
            var destWorkbook = new Workbook();

            WorksheetCollection sheets = destWorkbook.Worksheets;

            for (int i = 0; i < smallerCSVs.Length; i++)
            {
                var tempWorkbook = new Workbook(smallerCSVs[i]);

                sheets[i].Copy(tempWorkbook.Worksheets[0]);

                sheets.Add();
            }

            // save csv as ods
            destWorkbook.Save(outputODS, Aspose.Cells.SaveFormat.Auto);

            Console.WriteLine("Process ends.");
            Console.Read();
        }
    }
}
