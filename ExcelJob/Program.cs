using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelJob
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                GetValue();
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }


            Console.WriteLine("Please press any key to exit.");
            Console.ReadKey();
        }

        private static void GetValue()
        {
            string[] excelfiles = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.xlsx");
            for (int i = 0; i < excelfiles.Count(); i++)
            {
                string fileName = excelfiles[i];

                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                {
                    SharedStringTable sharedStringTable = document.WorkbookPart.SharedStringTablePart.SharedStringTable;
                    string cellValue = null;
                    string searchTarget = null;
                    string tValue = null;
                    string vValue = null;
                    string tMatchTemplate = "C-GTTRIBURF2";
                    string vMatchTemplate = "C-GTTRIBURF3";
                    string tMatch = tMatchTemplate + @"\s*(-?\d+(?:\.\d+)?)?";
                    string vMatch = vMatchTemplate + @"\s*(-?\d+(?:\.\d+)?)?";

                    foreach (WorksheetPart worksheetPart in document.WorkbookPart.WorksheetParts)
                    {
                        foreach (SheetData sheetData in worksheetPart.Worksheet.Elements<SheetData>())
                        {
                            if (sheetData.HasChildren)
                            {
                                foreach (Row row in sheetData.Elements<Row>())
                                {
                                    ArrayList hitTxts = new ArrayList();

                                    foreach (Cell cell in row.Elements<Cell>())
                                    {
                                        cellValue = cell.InnerText;
                                        var col = cell.CellReference.Value.Substring(0, 1);
                                        switch (col)
                                        {
                                            case "A":
                                                searchTarget = GetCellValue(cell, sharedStringTable, cellValue);

                                                string[] files = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.txt");
                                                for (int j = 0; j < files.Count(); j++)
                                                {
                                                    string fileContent = File.ReadAllText(files[j]);
                                                    if (fileContent.Contains(searchTarget))
                                                    {
                                                        hitTxts.Add(files[j]);
                                                    }
                                                }
                                                break;
                                            case "T":
                                                tValue = GetCellValue(cell, sharedStringTable, cellValue);
                                                break;
                                            case "V":
                                                vValue = GetCellValue(cell, sharedStringTable, cellValue);
                                                break;
                                        }
                                    }



                                    if (hitTxts != null && hitTxts.Count > 0)
                                    {
                                        foreach (string tFile in hitTxts)
                                        {
                                            Console.WriteLine(searchTarget + " in File: " + tFile);
                                            //do the replace
                                            string fileContent = File.ReadAllText(tFile);

                                            fileContent = ReplaceContent(fileContent, tMatch, tMatchTemplate, tValue);

                                            fileContent = ReplaceContent(fileContent, vMatch, vMatchTemplate, vValue);

                                            File.WriteAllText(tFile, fileContent);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    document.Close();
                }
            }
        }

        private static string GetCellValue(Cell cell, SharedStringTable sharedStringTable, string cellValue)
        {
            string returnString = null;

            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                returnString = sharedStringTable.ElementAt(Int32.Parse(cellValue)).InnerText;
                //Console.WriteLine("cell val: " + returnString);
            }
            else
            {
                returnString = cellValue;
                //Console.WriteLine("cell val: " + cellValue);
            }

            return returnString;
        }

        private static string ReplaceContent(string fileContent, string match, string matchTemplate, string val)
        {
            MatchCollection mc = Regex.Matches(fileContent, match);

            fileContent = Regex.Replace(fileContent, match, delegate (Match m) {
                return m.ToString().EndsWith("\r\n") ? matchTemplate + " " + val + "\r\n" : matchTemplate + " " + val;
            });

            //foreach (Match m in mc)
            //{
            //    Console.WriteLine(m);
            //    fileContent = fileContent.Replace(m.ToString().Replace("\r\n",""), matchTemplate + " " + val);
            //}

            return fileContent;
        }

    }
}
