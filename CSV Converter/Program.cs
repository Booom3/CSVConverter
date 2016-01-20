using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;

namespace CSV_Converter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Arg0 = Input file
            // Arg1 (Optional) = Output file
            string inputFile = "";
            string outputFile = "";
            if (args.Length > 0)
            {
                if (File.Exists(args[0]))
                {
                    inputFile = args[0];
                }
            }
            if (args.Length > 1)
            {
                if (File.Exists(args[1]))
                {
                    outputFile = args[1];
                }
            }
            if (inputFile == "")
                Environment.Exit(0);
            if (outputFile == "")
                outputFile = Path.Combine(Path.GetDirectoryName(inputFile),
                            Path.GetFileNameWithoutExtension(inputFile) + " output"
                            + Path.GetExtension(inputFile));
            Console.WriteLine("Input: " + inputFile);
            Console.WriteLine("Output: " + outputFile);
            // Read using CSV Helper
            var oldFile = new List<List<string>>();
            using (var reader = new StreamReader(inputFile))
            {
                var csvParser = new CsvParser(reader);
                while (true)
                {
                    var readLine = csvParser.Read();
                    if (readLine == null)
                        break;
                    oldFile.Add(readLine.ToList());
                }
            }

            // Specifically bring the first value of the specified columns down starting from 0
            // Probably shouldn't do this with columns that may be backslashed
            // But what do I know, I'm not the boss of you
            List<int> bringTheseColumnsDown = new List<int>();
            bringTheseColumnsDown.Add(15);
            bringTheseColumnsDown.Add(21);

            int totalColumns = oldFile.First().Count;
            // We'll just rebuild the file we want, it's easier than inserting into the old one
            var newFile = new List<string>();
            // First read row by row
            for (int row = 0; row < oldFile.Count; row++)
            {
                // The first value is an array of the split string
                // The second value is on which column the split belongs to
                var newLinesInColumn = new List<Tuple<string[], int>>();
                // How many new rows we need to insert
                int amountOfNewRows = 0;
                // Loop through all the columns on the row,
                // splitting them by forward slashes
                for (int column = 0; column < totalColumns; column++)
                {
                    string[] splitColumn = oldFile[row][column].Split('/');
                    // We only need to add as many new rows as the
                    // highest amount of splits
                    if (amountOfNewRows < splitColumn.Length)
                        amountOfNewRows = splitColumn.Length;

                    newLinesInColumn.Add(new Tuple<string[], int>(splitColumn, column));
                }
                var newLine = new StringBuilder();
                // Now we need to add the new rows
                for (int newRow = 0; newRow < amountOfNewRows; newRow++)
                {
                    for (int newColumn = 0; newColumn < totalColumns; newColumn++)
                    {
                        if (newLinesInColumn.ElementAt(newColumn).Item2 == newColumn)
                        {
                            if (newLinesInColumn.ElementAt(newColumn).Item1.Length > newRow)
                            {
                                newLine.Append(newLinesInColumn.ElementAt(newColumn).Item1[newRow]);
                            }
                            // Bring down any specific columns
                            else if (bringTheseColumnsDown.Contains(newColumn))
                            {
                                newLine.Append(newLinesInColumn.ElementAt(newColumn).Item1[0]);
                            }
                        }
                        newLine.Append(newColumn == totalColumns - 1 ? "\n" : ",");
                    }
                }
                newFile.Add(newLine.ToString());
            }
            
            using (var writer = new StreamWriter(outputFile))
            {
                foreach (var c in newFile)
                {
                    writer.Write(c);
                }
            }
            Console.WriteLine("Finished!");
            Console.ReadKey();
        }
    }
}
