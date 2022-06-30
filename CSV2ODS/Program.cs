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

            var directory = @"D:\CBS_Work\TeamOpenData\CSV2ODS\Data\";
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
            string[] inputCSVs = Directory.GetFiles(tempDir, "*.csv");

            string outputODS = Path.Combine(tempDir, "output.ods");

            // the final big output 
            var destWorkbook = new Workbook();
            
            for (int i = 0; i < inputCSVs.Length; i++)
            {
                WorksheetCollection sheets = destWorkbook.Worksheets;

                var tempWorkbook = new Workbook(inputCSVs[i]);

                sheets[i].Copy(tempWorkbook.Worksheets[0]);
                
                sheets.Add();

                Console.WriteLine(DateTime.Now + ": sheets" + i.ToString() + " is added to the workbook.");

                // save csv as ods
                destWorkbook.Save(outputODS, SaveFormat.Ods);
            }
            
            //// save csv as ods
            //destWorkbook.Save(outputODS, SaveFormat.Ods);

            Console.WriteLine("Process ends.");
            Console.Read();
        }
    }
}
