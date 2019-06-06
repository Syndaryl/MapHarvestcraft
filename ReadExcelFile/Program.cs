
using System;
using System.Collections.Generic;
using C = System.Console;
using ParseORE;
using CommandLineParser.Arguments;
using CommandLineParser.Exceptions;
using System.IO;

namespace ReadExcelFile
{
    class Program
    {

        static void Main(string[] args)
        {
            var range = new List<List<string>>();
            #region arguments
            // "Official Pam's HarvestCraft Food and Recipe List 1.12.x.xlsx" "Food 1.12.2zd" "B450" "O465"
            string excelFile="";
            string sheetName="";
            string firstCellName="";
            string lastCellName="";
            string oredictFile="";

            CommandLineParser.CommandLineParser commandLineParser = new CommandLineParser.CommandLineParser();
            ValueArgument<string> excelFilename = new ValueArgument<string>('e', "excel", "Excel spreadsheet");
            ValueArgument<string> worksheetName = new ValueArgument<string>('n', "name", "Name of the Worksheet to use from the given spreadsheet");
            ValueArgument<string> firstCell = new ValueArgument<string>('f', "first-cell", "Top-left cell to use from the given spreadsheet");
            ValueArgument<string> lastCell = new ValueArgument<string>('l', "last-cell", "Bottom-right cell  to use from the given spreadsheet");
            ValueArgument<string> oreDictFilename = new ValueArgument<string>('o',"ore-dictionary","Text file containing the OreDictionary dump from CraftTweaker");
            commandLineParser.Arguments.Add(excelFilename);
            commandLineParser.Arguments.Add(worksheetName);
            commandLineParser.Arguments.Add(firstCell);
            commandLineParser.Arguments.Add(lastCell);
            commandLineParser.Arguments.Add(oreDictFilename);

            try
            {
                commandLineParser.ParseCommandLine(args);
                if (excelFilename.Parsed)
                {

                    if (File.Exists(excelFilename.Value))
                    {
                        excelFile = excelFilename.Value;
                    } else
                    {
                        
                        throw new FileNotFoundException(string.Format("'{0}/{1}' does not exist, please check your filename and path.",Directory.GetCurrentDirectory(),excelFilename.Parsed));
                    }
                }
                if (oreDictFilename.Parsed)
                {
                    if (File.Exists(oreDictFilename.Value))
                    {
                        oredictFile = oreDictFilename.Value;
                    }
                    else
                    {
                        throw new FileNotFoundException(string.Format("'{0}' does not exist, please check your filename and path.", excelFilename.Parsed));
                    }
                }
                if (worksheetName.Parsed)
                {
                    sheetName = worksheetName.Value;
                }
                if (firstCell.Parsed)
                {
                    firstCellName = firstCell.Value;
                }
                if (lastCell.Parsed)
                {
                    lastCellName = lastCell.Value;
                }
            } catch (CommandLineException e)
            {
                Console.WriteLine(e.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            #endregion



            if (!string.IsNullOrEmpty(firstCellName) && !string.IsNullOrEmpty(lastCellName))
            {
                range = ExcelRangeFromSheet(excelFile, sheetName, firstCellName, lastCellName);
            }
            else
            {
                range = ExcelSheet(excelFile, sheetName);
            }

            Dictionary<string, Item> items = ItemDictionary(cellRange: range, column: 3);

            var parser = new OREParser(oredictFile, System.Text.Encoding.UTF8, items: items);
            parser.ReadAllOres();
            //DumpOres(parser);

            var recipes = new List<BenchFoodRecipe>();
            foreach (var row in range)
            {
                if (!"Amount".Equals(row[5]) )
                    recipes.Add(BenchFoodRecipe.BenchFoodRecipeFactory(recipes: recipes, row: row, oreDictionary: parser.Ores, items: items));
            }

            C.WriteLine("Press <Enter> to quit");
            C.ReadLine();
        }

        private static Dictionary<string, Item> ItemDictionary(List<List<string>> cellRange, int column = 3)
        {
            Dictionary<string, Item> result = new Dictionary<string, Item>(cellRange.Count);
            foreach (var row in cellRange)
            {
                if (! string.IsNullOrEmpty( row[column]))
                {
                    result[row[column]] = new Item(row[column]);
                }
            }


            return result;
        }

        private static void DumpOres(OREParser parser)
        {
            C.WriteLine("There are {0} ore dictionary categories.", parser.Ores.Count);
            foreach (var oDictEntry in parser.Ores)
            {
                C.WriteLine("{0} contains {1} synonyms", oDictEntry.Value.Name, oDictEntry.Value.Count);
            }
        }

        private static List<List<string>> ExcelSheet(string inputFilename, string sheetName)
        {
            List<List<string>> range;
            using (var handler = new WorkSheetHandler())
            {
                C.WriteLine("Loading from Excel file, please wait...");
                handler.SetExcelDocument(inputFilename);
                handler.On100thRow += delegate (object sender, EventArgs e)
                {
                    C.Write(".");
                };
                range = handler.Sheet(sheetName, true);
                C.WriteLine();
                C.WriteLine("Excel file loaded.");
            }

            return range;

        }

        private static List<List<string>> ExcelRangeFromSheet(string inputFilename, string sheetName, string firstCellName, string lastCellName)
        {
            List<List<string>> range;
            using (var handler = new WorkSheetHandler())
            {
                C.WriteLine("Loading from Excel file, please wait...");
                // handler.DumpSheetRange(C.Out, inputFilename, sheetName, firstCellName, lastCellName);
                handler.SetExcelDocument(inputFilename);
                handler.On100thRow += delegate (object sender, EventArgs e)
                {
                    C.Write(".");
                };
                range = handler.SheetRange(sheetName, firstCellName, lastCellName);
                C.WriteLine();
                C.WriteLine("Excel file loaded.");
            }

            return range;
        }
    }
}
