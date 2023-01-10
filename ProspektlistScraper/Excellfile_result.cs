using System;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ProspektlistScraper
{
    static class Excellfile_result
    {
        static Application excelApp;
        static Workbook excelWorkbook;
        static Worksheet excelWorksheet;

        static int row = 0;
        static public void Create()
        {
            row = 0;

            if (excelApp == null)
            {
                //Kasta en exception om det inte funkar att skapa en fil
                excelApp = new Application();
                if (excelApp == null)
                {
                    throw new Exception("Det gick inte att skapa excel filen .");
                }
            }
            if(null==excelWorkbook)
                excelWorkbook = excelApp.Workbooks.Add();
            
            if(null==excelWorksheet)
                excelWorksheet = (Worksheet)excelWorkbook.Sheets.Add();

            //skapa header i första raden
            excelWorksheet.Cells[1, 1] = "Företagsnamn";
            excelWorksheet.Cells[1, 2] = "Hemsida";
            excelWorksheet.Cells[1, 3] = "Org.nummer";
            excelWorksheet.Cells[1, 4] = "GTM";
            excelWorksheet.Cells[1, 5] = "UA";
            excelWorksheet.Cells[1, 6] = "H1";
            excelWorksheet.Cells[1, 7] = "H2";

            //set backgundsfärg för header rad
            excelWorksheet.Cells[1, 1].Interior.Color = XlRgbColor.rgbSilver;
            excelWorksheet.Cells[1, 2].Interior.Color = XlRgbColor.rgbSilver;
            excelWorksheet.Cells[1, 3].Interior.Color = XlRgbColor.rgbSilver;
            excelWorksheet.Cells[1, 4].Interior.Color = XlRgbColor.rgbSilver;
            excelWorksheet.Cells[1, 5].Interior.Color = XlRgbColor.rgbSilver;
            excelWorksheet.Cells[1, 6].Interior.Color = XlRgbColor.rgbSilver;
            excelWorksheet.Cells[1, 7].Interior.Color = XlRgbColor.rgbSilver;

            //ge varje kolumn en fördefinierat bredd
            excelWorksheet.Columns[1].ColumnWidth = 25;
            excelWorksheet.Columns[2].ColumnWidth = 25;
            excelWorksheet.Columns[3].ColumnWidth = 12;
            excelWorksheet.Columns[4].ColumnWidth = 15;
            excelWorksheet.Columns[5].ColumnWidth = 15;
            excelWorksheet.Columns[6].ColumnWidth = 60;
            excelWorksheet.Columns[7].ColumnWidth = 60;

            int counterForFile = 0;

            //kolla om fil med samma namn existera, ifall att sätt counterForFile +1
            //for (counterForFile = 0; File.Exists(@"C:\SkrapResultatLista" + counterForFile + ".xls"); counterForFile++)
            for (counterForFile = 0; File.Exists(@"C:\SkrapResultatLista" + " " + DateTime.Now.ToString("ddMMyyyy") + " " + counterForFile + ".xls"); counterForFile++)
            {

            }
            excelApp.ActiveWorkbook.SaveAs(@"C:\SkrapResultatLista" + " " + DateTime.Now.ToString("ddMMyyyy") + " " + counterForFile + ".xls", XlFileFormat.xlWorkbookNormal);
        }

        static public void Write(string[] scrapItem, string companyname, string homepage)
        {
            row++;

            //första item är n+1 för att excel börjar räkna från 1 och inte från 0 plus att vi har en header i första raden
            excelWorksheet.Cells[row + 1, 1] = companyname;
            excelWorksheet.Cells[row + 1, 2] = homepage;
            excelWorksheet.Cells[row + 1, 3] = scrapItem[0];
            excelWorksheet.Cells[row + 1, 4] = scrapItem[1];
            excelWorksheet.Cells[row + 1, 5] = scrapItem[2];
            excelWorksheet.Cells[row + 1, 6] = scrapItem[3];
            excelWorksheet.Cells[row + 1, 7] = scrapItem[4];

            //set texten i varje cell så den är längst upp i cellen
            excelWorksheet.Cells[row + 1, 1].VerticalAlignment = XlVAlign.xlVAlignTop;
            excelWorksheet.Cells[row + 1, 2].VerticalAlignment = XlVAlign.xlVAlignTop;
            excelWorksheet.Cells[row + 1, 3].VerticalAlignment = XlVAlign.xlVAlignTop;
            excelWorksheet.Cells[row + 1, 4].VerticalAlignment = XlVAlign.xlVAlignTop;
            excelWorksheet.Cells[row + 1, 5].VerticalAlignment = XlVAlign.xlVAlignTop;
            excelWorksheet.Cells[row + 1, 6].VerticalAlignment = XlVAlign.xlVAlignTop;
            excelWorksheet.Cells[row + 1, 7].VerticalAlignment = XlVAlign.xlVAlignTop;

            excelApp.ActiveWorkbook.Save();
        }

        static public void Close()
        {
            excelWorkbook.Close();
            excelApp.Quit();

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
            
            excelApp = null;
            excelWorkbook = null;
            excelWorksheet = null;
        }
    }
}
