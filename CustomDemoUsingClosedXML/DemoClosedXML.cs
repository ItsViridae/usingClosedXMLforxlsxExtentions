using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CustomDemoUsingClosedXML
{
    public class DemoClosedXML
    {
        public void GetExcelFile()
        {

            string fileName = @"C:\Users\scomeaux\source\repos\CustomDemoUsingClosedXML\CustomDemoUsingClosedXML\DemoExcelData.xlsx"; 


           using (var workbook = new XLWorkbook(fileName))
           {
                var ws1 = workbook.Worksheets.Worksheet(1); //FIRST WorkSheet
                ConvertWorksheetTo2dArray(ws1); //passes current worksheet(1)

            }

        }

        public string[,] ConvertWorksheetTo2dArray(IXLWorksheet worksheet)
        {
            // Look for the first row used
            var firstRowUsed = worksheet.FirstRowUsed();

            // Narrow down the row so that it only includes the used part
            var residentinfoHeaders = firstRowUsed.RowUsed();
            var residentinfo = residentinfoHeaders;

            Console.WriteLine($"Show Resident Information: ", residentinfoHeaders.Cell(1).Value);

            var i = 1;
            while (!residentinfo.Cell(i).IsEmpty())
            {
                //Console.WriteLine();
                Console.WriteLine(residentinfo.Cell(i).Value + "|");
                i++;

            }

           

            // for(InitializeSomething; BooleanCheckSomething; MutateSomething(s))
            //
            // 

            for (var row = firstRowUsed.RowUsed(); !row.IsEmpty(); row = row.RowBelow())
            {
                foreach (var cell in row.Cells())
                {
                    Console.Write(cell.Value);
                    Console.Write(" ");
                }
                Console.WriteLine();
            }


            //Index's for The Cells in Table
            var activeRowCount = worksheet.RowsUsed().Count();
            var activeColumnCount = worksheet.ColumnsUsed().Count();

            string[,] my2dArrayWithConvertedCells = new string[,]
            {
                {"",""},
            };
            

            //using foreach to create 2d Array
            //convert to 2d Array and return the values
            var rangeOfFirstNames = worksheet.Range("A1:A5");
            foreach (var cell in rangeOfFirstNames.Cells())
            {
                cell.SetDataType(XLDataType.Text);
                var valueofcell = cell.Value.ToString();

            }
            return my2dArrayWithConvertedCells;
        }
    }
}
