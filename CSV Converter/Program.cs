using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using CsvHelper;
using System.Xml.Linq;

// This program takes a CSV file in the shape of an SQL dump and explodes the forward slashes
// into their own rows. So for example a CSV file looking like:
// AAA,BBB,LLL/MMM/NNN,CCC
// Would turn into a CSV file looking like:
// AAA,BBB,LLL/MMM/NNN,CCC
// ,,LLL,
// ,,MMM,
// ,,NNN,
// And when you input that into excel:
// [AAA][BBB][LLL/MMM/NNN][CCC]
// [   ][   ][LLL]        [   ]
// [   ][   ][MMM]        [   ]
// [   ][   ][NNN]        [   ]
//
// It can also bring down a column with each exploded forward slash
// Say we want to bring down column 2 (array position 1),
// after we're done in excel it would look like:
// [AAA][BBB][LLL/MMM/NNN][CCC]
// [   ][BBB][LLL]        [   ]
// [   ][BBB][MMM]        [   ]
// [   ][BBB][NNN]        [   ]

namespace CSV_Converter
{
    class Program
    {
        static string appDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CSVConverter");
        static string configFile = "config.xml";
        static string configFilePath = Path.Combine(appDataFolder, configFile);
        static void CreateConfig()
        {
            if (!Directory.Exists(appDataFolder))
            {
                Directory.CreateDirectory(appDataFolder);
            }
            var config =
                new XDocument(new XElement("CSVConversion",
                    new XComment("All Column tags accept both numbers and letters. For example you can use Column 0 or Column A, Column 36 or Column VP. They are the same."),
                    new XElement("BringDownColumns",
                        new XComment("Add one element for each column you want to bring down. -1 does nothing."),
                        new XElement("Column", -1)
                    ),
                    new XElement("ExplodeColumns",
                        new XComment("Add one element for each column you want to explode. -1 does nothing."),
                        new XElement("Column", -1)
                    ),
                    new XComment("Sets the output directories. Supports multiple paths and environment variables. Empty means put one in the same directory as the input."),
                    new XElement("OutputDirectory",
                        new XElement("Path", "")
                    )
                ));
            config.Save(configFilePath);
        }

        static int ConfigParseColumn(string input)
        {
            int ret = 0;
            if (int.TryParse(input, out ret))
                return ret;
            else
            {
                for (int i = 0; i < input.Length; i++)
                {
                    // Alphabet position - 1 (converted to array position) (case insensitive)
                    // This is the column it's labeled in Excel
                    ret += ((int)input[i] % 32) - 1;
                }
                return ret;
            }
        }

        static void Main(string[] args)
        {
            // Arg0 = Input file
            string inputFile = "";
            List<string> outputFile = new List<string>();
            if (args.Length > 0)
            {
                if (File.Exists(args[0]))
                {
                    inputFile = args[0];
                }
            }

            if (inputFile == "")
                Environment.Exit(0);

            XDocument config;
            List<int> bringTheseColumnsDown = new List<int>();
            List<int> explodeTheseColumns = new List<int>();
            if (!File.Exists(configFilePath))
            {
                Console.WriteLine("New Config");
                CreateConfig();
            }
            config = XDocument.Load(configFilePath);
            var bringDowns =
                from el in config.Root.Element("BringDownColumns").Elements("Column")
                select el;
            bringTheseColumnsDown = bringDowns.Select(x => ConfigParseColumn(x.Value)).ToList();
            var explodeLines =
                from el in config.Root.Element("ExplodeColumns").Elements("Column")
                select el;
            explodeTheseColumns = explodeLines.Select(x => ConfigParseColumn(x.Value)).ToList();
            var outputPath =
                from el in config.Root.Element("OutputDirectory").Elements("Path")
                select el;
            
            foreach (var p in outputPath)
            {
                string path = Environment.ExpandEnvironmentVariables(p.Value);
                outputFile.Add(Path.Combine((path == "" ? Path.GetDirectoryName(inputFile) : path),
                            Path.GetFileNameWithoutExtension(inputFile) + " output"
                            + Path.GetExtension(inputFile)));
            }
            Console.WriteLine("Input: " + inputFile);
            foreach (var b in outputFile)
                Console.WriteLine("Output: " + b);
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
            
            int totalColumns = oldFile.First().Count;
            // We'll just rebuild the file we want, it's easier than inserting into the old one
            var newFile = new List<string>();
            // First read row by row
            for (int row = 0; row < oldFile.Count; row++)
            {
                // Preserve each row of the old file
                newFile.Add(string.Join(",", oldFile[row]) + "\n");
                // We explode each string into an array, and we put this in a list
                // where each entry in the list is one column
                var newLinesInColumn = new List<string[]>();
                // How many new rows we need to insert
                int amountOfNewRows = 0;
                // Loop through all the columns on the row,
                // splitting them by forward slashes
                for (int column = 0; column < totalColumns; column++)
                {
                    if (explodeTheseColumns.Contains(column))
                    {
                        string[] splitColumn = oldFile[row][column].Split('/');
                        // We only need to add as many new rows as the
                        // highest amount of splits
                        if (amountOfNewRows < splitColumn.Length)
                            amountOfNewRows = splitColumn.Length;

                        newLinesInColumn.Add(splitColumn);
                    }
                    else if (bringTheseColumnsDown.Contains(column))
                    {
                        newLinesInColumn.Add(new string[] { oldFile[row][column] });
                    }
                    else
                    {
                        newLinesInColumn.Add(new string[]{ });
                    }
                }
                var newLine = new StringBuilder();
                // Now we need to add the new rows
                // Only add a new row if we have more than 1 to add, the old row has already been
                // added so this should only be 2 or more when a string was exploded
                if (amountOfNewRows > 1)
                {
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
                }
                newFile.Add(newLine.ToString());
            }

            foreach (var o in outputFile)
            {
                using (var writer = new StreamWriter(o))
                {
                    foreach (var c in newFile)
                    {
                        writer.Write(c);
                    }
                }
            }
            Console.WriteLine("Finished!");
            Console.ReadKey();
        }
    }
}
