using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Drawing;

namespace ProspektlistScraper
{
    public partial class Form1 : Form
    {
        enum ThreadState
        {
            inactive, running, canceling
        };

        ThreadState threadState;     //variabel for ending all threads
        SitemapGenerator.Sitemap.Sitemapper siteMapper;
        public Form1()
        {
            InitializeComponent();

            siteMapper = null;

            progressBar1.Maximum = 0;
            progressBar1.Minimum = 0;
            progressBar1.Step = 1;
            progressBar1.Style = ProgressBarStyle.Blocks;
            progressBar1.Value = 0;
            //progressBar1.ForeColor = Color.Red;

            this.FormClosing += Form1_FormClosing;
        }

        /// <summary>
        /// Override the Close Form event
        /// Do something
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_FormClosing(Object sender, FormClosingEventArgs e)
        {
            if(threadState == ThreadState.inactive)
            {
                return;
            }
            // Assume that X has been clicked and act accordingly.
            // Confirm user wants to close
            if (DialogResult.No == MessageBox.Show(this, "Avbrytdialog" , "Är du säker att du vill avbryta?", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                //Stay on this form
                e.Cancel = true;
            }
            else
            {
                if (threadState != ThreadState.inactive)
                {
                    //End all threads because of user canceling
                    threadState = ThreadState.canceling;

                    label2.Text = "AVBRYTER!!!!";
                    label2.ForeColor = Color.Red;

                    if (siteMapper != null)
                    {
                        siteMapper.endThread = true;
                    }

                    //waiting explicitly for status canceling
                    while (ThreadState.canceling == threadState)
                    {
                        Thread.Sleep(500);
                    }
                }
            }                            
        }
       
        private void selectFileButton_Click(object sender, EventArgs e)
        {
            try
            {
                //Öppna dialogfönster för att kunna välja ut önskat fil
                using (OpenFileDialog fil = new OpenFileDialog())
                {
                    if(DialogResult.OK == fil.ShowDialog())
                    {
                        //Visa medelandet om börjana av skrapning med möjlighet att avbryta
                        string message = "Appen börjar skrapa nu. Tack för ditt tålamod!";
                        string title = "Börja skrapa!";
                        MessageBoxButtons buttons = MessageBoxButtons.OKCancel;
                        
                        //this is for making dialog child (modal)
                        DialogResult resultStart = MessageBox.Show(this,message, title, buttons);

                        if (resultStart == DialogResult.OK)
                        {
                            Thread t = new Thread((ThreadStart)(() =>
                            {

                                //Ta med filsökväg till funktionen som öppnar och bearbetar filen
                                ExcelFileReader(fil.FileName.ToString());

                                //signal for end-dialog-event (X) that all threads are closed
                                threadState = ThreadState.inactive;

                                //the program uses MS Excel via OLE. Set memory of OLE free.
                                GC.Collect();
                                GC.WaitForPendingFinalizers();

                                MessageBox.Show("Skrapning är klart! Hämta din scrapresultlista här: \nC:\\SkrapResultatLista ddmmyyyy 0.xls!\nListan med felaktiga url ligger under: \nC:\\IncorrectUrlLista ddmmyyyy 0.xls!");

                            }));

                            //t.SetApartmentState(ApartmentState.STA);
                            threadState = ThreadState.running;
                            t.Start();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void ExcelFileReader(string path)
        {
            try
            {
                selectFileButton.BeginInvoke(new MethodInvoker(delegate { selectFileButton.Enabled = false; }));

                // ExcelReaderFactory needs this line for knowing codepage 1252
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                //öppna förvalt excel fil
                var stream = File.Open(path, FileMode.Open, FileAccess.Read);
                var reader = ExcelReaderFactory.CreateReader(stream);
                var result = reader.AsDataSet();
                var tables = result.Tables.Cast<DataTable>();

                stream.Close();
                reader.Close();

                //skapa lista för alla url som hittades i excel filen
                List<string> urlList = new List<string>();
                
                //skapa lista för alla filnamn för tillhörande url från excel filen
                List<string> filnamnList = new List<string>();

                //incorectUrlList = innehåller alla url som inte var valid
                List<string> incorrectUrlList = new List<string>();

                string webpage = "";

                foreach (DataTable table in tables)
                {
                    foreach (DataRow row in table.Rows)
                    {
                        //urlListLine = innehål i tredje kolumn (företags url) i vald excel fil
                        string urlListLine = row.ItemArray[2].ToString().ToLower().Trim();
                        
                        //filename = innehål i andra kolumn (företagsnamn) i vald excel fil
                        string filename = row.ItemArray[1].ToString().ToLower().Trim();

                        //kontroll om url har en giltig längd och börjar med https:// eller www. I fall att den bara börja med www sätt https framför.
                        if ((8 <= urlListLine.Length && urlListLine.Substring(0, 8) == "https://"))
                        {
                            //take the url and check if valid
                            if (ValidateUrl(urlListLine) == true)
                            {
                                webpage = urlListLine;
                                urlList.Add(webpage);

                                filnamnList.Add(filename);
                            }
                            else
                            {
                                incorrectUrlList.Add(urlListLine);
                            }
                                    
                        }
                        else if ((4 <= urlListLine.Length && urlListLine.Substring(0, 4) == "www.") && (urlListLine.Substring(0, 8) != "https://"))
                        {
                            string korrekturlListLine = "https://" + urlListLine;

                            //take the url and check if valid
                            if (ValidateUrl(korrekturlListLine) == true)
                            {
                                webpage = korrekturlListLine;
                                urlList.Add(webpage);

                                filnamnList.Add(filename);
                            }
                            else
                            {
                                incorrectUrlList.Add(korrekturlListLine);
                            }
                        }
                    }
                }

                ScraperAsync(urlList, filnamnList, incorrectUrlList);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Function to validate URL
        // using regular expression
        public static bool ValidateUrl(string url)
        {
            return Uri.IsWellFormedUriString(url, UriKind.Absolute);
        }

        public void ScraperAsync(List<string> urlList, List<string> filenameList, List<string> incorrectUrlList)
        {
                try
                {
                    Excellfile_result.Create();

                    int n = 0;

                    //counts amount rows to scrape
                    int sumAvRows = urlList.Count();
                    progressBar1.BeginInvoke(new MethodInvoker(delegate { progressBar1.Maximum = sumAvRows; }));
                    label2.BeginInvoke(new MethodInvoker(delegate { label2.Text = progressBar1.Value.ToString() + " / " + sumAvRows.ToString(); }));

                    //List som sammlar ihop alla scrapresult
                    List<string[]> scrapresultList = new List<string[]>();

                    string domain = "";

                    foreach (string url in urlList)
                    {
                        //vi provar om domänsuffix är korrekt
                        if (12 > url.Length)
                        {
                            throw new Exception("Your webadress is too short. It must be at least 12 characters long.");
                        }

                        //vi går utifrån at det alltid finns https://wwww.
                        string[] domainSplit = url.Split('.');
                        string domain_prefix = domainSplit[1];

                        string domain_suffix = "";
                        try
                        {
                            for (int suffixIndex = 0; suffixIndex < domainSplit[2].Length; suffixIndex++)
                            {
                                if ('a' > domainSplit[2][suffixIndex] || 'z' < domainSplit[2][suffixIndex])
                                    break;

                                domain_suffix += domainSplit[2][suffixIndex];
                            }
                        }
                        catch (Exception ex)
                        {
                            incorrectUrlList.Add(url);
                            //MessageBox.Show(ex.Message);
                        }

                        domain = domain_prefix + "." + domain_suffix;

                        siteMapper = new SitemapGenerator.Sitemap.Sitemapper(domain, url);
                    
                        bool done = false;
                        Thread t = new Thread((ThreadStart)(async () =>
                        {
                            await siteMapper.GenerateSitemap();
                            done = true;
                        }));

                        t.SetApartmentState(System.Threading.ApartmentState.STA);
                        t.Start();

                        t.Join();
                        while (!done) { };

                        string[] myScrapItem = new string[5];

                        myScrapItem[0] = siteMapper.orgString;
                        myScrapItem[1] = siteMapper.gtmString;
                        myScrapItem[2] = siteMapper.uaString;
                        myScrapItem[3] = siteMapper.GetUniqueH1();
                        myScrapItem[4] = siteMapper.GetUniqueH2();

                        scrapresultList.Add(myScrapItem);

                        progressBar1.BeginInvoke(new MethodInvoker(delegate { progressBar1.PerformStep(); }));
                        label2.BeginInvoke(new MethodInvoker(delegate { label2.Text = progressBar1.Value.ToString() + " / " + sumAvRows.ToString(); }));

                        //putt all the results one by one into file

                        if(threadState != ThreadState.canceling)
                        {
                        Excellfile_result.Write(myScrapItem, filenameList[n], urlList[n]);
                        
                        }
                        n++;

                        siteMapper = null;

                        if (threadState == ThreadState.canceling)
                        {
                            break;
                        }
                    }

                    //Write to file
                    FileWriterIncorrectUrllist(incorrectUrlList);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    Excellfile_result.Close();

                    progressBar1.BeginInvoke(new MethodInvoker(delegate { progressBar1.Value = 0; }));
                    selectFileButton.BeginInvoke(new MethodInvoker(delegate { selectFileButton.Enabled = true; }));
                }
        }

        public void FileWriterIncorrectUrllist(List<string> incorrectUrlList)
        {
            int n = 0;

            if(incorrectUrlList.Count == 0)
            {
                return;
            }

            //Kasta en exception om det inte funkar att skapa en fil
            Excel.Application excelApp = new Excel.Application();
            if (excelApp == null)
            {
                throw new Exception("Det gick inte att skapa excel filen .");
            }

            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets.Add();
            //skapa header i första raden
            excelWorksheet.Cells[1, 1] = "Felaktig url";

            //set backgundsfärg för header rad
            excelWorksheet.Cells[1, 1].Interior.Color = Excel.XlRgbColor.rgbSilver;

            //gå igenom lista med alla skrapade resultat
            foreach (string felUrl in incorrectUrlList)
            {
                n++;

                //första item är n+1 för att excel börjar räkna från 1 och inte från 0 plus att vi har en header i första raden
                excelWorksheet.Cells[n + 1, 1] = felUrl;

                //set texten i varje cell så den är längst upp i cellen
                excelWorksheet.Cells[n + 1, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            }

            excelWorksheet.Columns[1].ColumnWidth = 25;

            int counterForFile = 0;

            //kolla om fil med samma namn existera, ifall att sätt counterForFile +1
            for (counterForFile = 0; File.Exists(@"C:\IncorrectUrlLista" + " " + DateTime.Now.ToString("ddMMyyyy") + " " + counterForFile + ".xls"); counterForFile++)
            {

            }
            excelApp.ActiveWorkbook.SaveAs(@"C:\IncorrectUrlLista" + " " + DateTime.Now.ToString("ddMMyyyy") + " " + counterForFile + ".xls", Excel.XlFileFormat.xlWorkbookNormal);

            excelWorkbook.Close();
            excelApp.Quit();

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);

        }
    }
}