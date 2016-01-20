using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;

// This program takes a CSV file in the shape of an SQL dump and explodes the forward slashes
// into their own rows. So for example a CSV file looking like:
// AAA,BBB,LLL/MMM/NNN,CCC
// Would turn into a CSV file looking like:
// AAA,BBB,LLL,CCC
// ,,MMM,
// ,,NNN,
// And when you input that into excel:
// [AAA][BBB][LLL][CCC]
// [   ][   ][MMM][   ]
// [   ][   ][NNN][   ]
//
// It can also bring down a column with each exploded forward slash
// Say we want to bring down column 2 (array position 1),
// after we're done in excel it would look like:
// [AAA][BBB][LLL][CCC]
// [   ][BBB][MMM][   ]
// [   ][BBB][NNN][   ]

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
            // Probably shouldn't do this with columns that may be forward slashed
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
                // We explode each string into an array, and we put this in a list
                // where each entry in the list is one column
                var newLinesInColumn = new List<string[]>();
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

                    newLinesInColumn.Add(splitColumn);
                }
                var newLine = new StringBuilder();
                // Now we need to add the new rows
                for (int newRow = 0; newRow < amountOfNewRows; newRow++)
                {
                    // Make sure we get the same amount of columns
                    for (int newColumn = 0; newColumn < totalColumns; newColumn++)
                    {
                        // If there is something in the column to insert, insert it
                        if (newLinesInColumn[newColumn].Length > newRow)
                        {
                            // Add the exploded data into the new column
                            newLine.Append(newLinesInColumn[newColumn][newRow]);
                        }
                        // If there is no exploded data to insert but we want to bring this column down
                        else if (bringTheseColumnsDown.Contains(newColumn))
                        {
                            // Insert the first value into the column, which is always
                            // the same whether it was exploded or not
                            newLine.Append(newLinesInColumn[newColumn][0]);
                        }
                        // If it's the last entry on the line, just put newline
                        // Otherwise, whether we put something on this line or not, put a comma
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
