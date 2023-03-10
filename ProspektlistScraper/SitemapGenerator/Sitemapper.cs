using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;

namespace SitemapGenerator.Sitemap
{
    public class Sitemapper
    {
        public SitemapDocument Document;
        public string BaseUrl { get; set; }
        public string Domain { get; set; }
        public string SavePath { get; set; }
        public bool Exclude { get; set; }
        private ILoader _loader = new Loader();
        public delegate void Info();
        //public Info Notify;
        private int timeSetter = 1; //default = 1

        public bool endThread = false;

        public Sitemapper(string domain, string baseulr)
        {
            endThread = false;

            Domain = domain;
            BaseUrl = baseulr.Contains("http://") || baseulr.Contains("https://") ? baseulr : ("https://" + baseulr);
            SavePath = DateTime.Now.ToString("yyyymmddhhmmss") + ".xml";
            Exclude = true;
            Document = new SitemapDocument();
        }

        public async Task GenerateSitemap()
        {
            try
            {
                List<string> new_urls = new List<string>();
                List<string> visited = new List<string>();
                Stopwatch sw = new Stopwatch();

                listDataH1 = new List<string>();
                listDataH2 = new List<string>();
                Document.Urls = visited;
                new_urls.Add(BaseUrl);

                sw.Start();

                do
                {
                    List<string> hrefs = new List<string>();

                    foreach (var url in new_urls)
                    {
                        string text = "";

                        text = await _loader.Get(url);
                        if (string.IsNullOrEmpty(text)) continue;

                        visited.Add(url);

                        //Suche nach Skrapkriterien
                        await GetData(url);

                        if (sw.Elapsed.TotalMinutes > timeSetter || endThread == true)
                        {
                            break;
                        }

                        //Notify?.Invoke();

                        List<string> meta = new List<string>();

                        try
                        {
                            /*List<string>*/ meta = Parser.GetAHrefs(text).Distinct().ToList();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                        meta = HrefFilter(meta);

                        Parser.Normalize(Domain, url, ref meta);

                        List<string> meta2 = new List<string>();

                        meta2 = meta.Select(u => u.Contains('?') ? u.Split('?')[0] : u).ToList();

                        if (Exclude)
                            meta = meta.Select(u => u.Contains('?') ? u.Split('?')[0] : u).ToList();

                        hrefs.AddRange(meta);
                        hrefs = hrefs.Distinct().ToList();

                    }
                    new_urls = hrefs.Except(visited).ToList();
                }
                while (new_urls.Count != 0 && sw.Elapsed.TotalMinutes <= timeSetter && endThread == false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private List<string> HrefFilter(List<string> metaList)
        {
            try
            {
                List<string> toRemove = new List<string>();

                foreach (string meta in metaList)
                {
                    //search for: www. https: http:
                    if (4 > meta.Length)
                    {
                        toRemove.Add(meta);
                    }
                    else
                    {
                        if (meta.Substring(0, 4).ToLower() != "www.")
                        {
                            if (5 > meta.Length)
                            {
                                toRemove.Add(meta);
                            }
                            else
                            {
                                if (meta.Substring(0, 5).ToLower() != "http:")
                                {
                                    if (6 > meta.Length)
                                    {
                                        toRemove.Add(meta);
                                    }
                                    else
                                    {
                                        if (meta.Substring(0, 6).ToLower() != "https:")
                                        {
                                            toRemove.Add(meta);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                foreach (string removeItem in toRemove)
                {
                    metaList.Remove(removeItem);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return metaList;
        }

        //kod for web-scraping/web-parsing

        public string gtmString = "";
        public string uaString = "";
        List<string> listDataH1;
        List<string> listDataH2;
        public string orgString = "";

        public string GetUniqueH1()
        {
            string uniqueH1String = "";
                foreach (string item in listDataH1.Distinct())
                {
                    uniqueH1String += item + "\n";
                }

            return uniqueH1String;
        }

        public string GetUniqueH2()
        {
            string uniqueH2String = "";

                foreach (string item in listDataH2.Distinct())
                {
                    uniqueH2String += item + "\n";
                }

            return uniqueH2String;
        }
        
        public async Task GetData(string url)
        {
            try
            {
                var httpClient = new HttpClient();

                var htmlDocument = new HtmlAgilityPack.HtmlDocument();

                var html = "";
                    
                html = await httpClient.GetStringAsync(url);
                htmlDocument.LoadHtml(html);


                //Try to get Orgnr from website
                HtmlNode orgNode = htmlDocument.DocumentNode.SelectSingleNode("//*[text()[contains(.,'Org')]]");

                if (null != orgNode)
                {
                    string nodeResult = orgNode.InnerText.ToUpper();

                    int orgIndex = nodeResult.IndexOf("ORG");
                    while (orgIndex < nodeResult.Length && !char.IsDigit(nodeResult[orgIndex]))
                    {
                        orgIndex++;
                    }
                    if (nodeResult.Length > orgIndex)
                    {
                        if (10 < nodeResult.Length - orgIndex)
                        {

                            string orgResult = nodeResult.Substring(orgIndex, 11);

                            (bool result, string strOrg) = OrganisationsNummer(orgResult);

                            if (true == result)
                            {
                                orgString = strOrg;
                            }
                        }
                    }
                }

                //Try to get GTM data from website
                HtmlNode gtmNode = htmlDocument.DocumentNode.SelectSingleNode("//*[text()[contains(., 'GTM-')]]");

                if (null != gtmNode)
                {
                    string gtmResult = gtmNode.InnerText;

                    int gtmIndex = gtmResult.IndexOf("GTM-");

                    if (gtmIndex >= 0)
                    {
                        int gtmEndIndex = gtmResult.IndexOf('\'', gtmIndex);

                        if (gtmEndIndex >= 0)
                        {
                            gtmString = gtmResult.Substring(gtmIndex, gtmEndIndex - gtmIndex);
                        }
                    }
                }

                //Try to get GA(UA number) data from website
                HtmlNode node = htmlDocument.DocumentNode.SelectSingleNode("//*[text()[contains(., 'UA-')]]");

                if (null != node)
                {
                    string uaResult = node.InnerText;

                    int uaIndex = uaResult.IndexOf("UA-");

                    if (uaIndex >= 0)
                    {
                        int uaEndIndex = uaResult.IndexOf('\'', uaIndex);

                        if (uaEndIndex >= 0)
                        {
                            if (uaEndIndex - uaIndex > 2 && "-1" == uaResult.Substring(uaEndIndex - 2, 2))
                            {
                                uaString = uaResult.Substring(uaIndex, uaEndIndex - uaIndex);
                            }
                        }
                    }
                }

                //Try to get H1 data from website
                HtmlNodeCollection h1NodeList = htmlDocument.DocumentNode.SelectNodes($"//h1");

                if (null != h1NodeList)
                {
                    foreach (var item in h1NodeList)
                    {
                        string trimmed = item.InnerText.Trim();

                        if (0 < trimmed.Length)
                        {
                            listDataH1.Add(trimmed);
                        }
                    }
                }

                //Try to get H2 data from website
                HtmlNodeCollection h2NodeList = htmlDocument.DocumentNode.SelectNodes($"//h2");

                if (null != h2NodeList)
                {
                    foreach (var item in h2NodeList)
                    {
                        string trimmed = item.InnerText.Trim();

                        if (0 < trimmed.Length)
                        {
                            listDataH2.Add(trimmed);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }


        static (bool, string) OrganisationsNummer(string nr)
        {

            Regex rg = new Regex(@"(\d{1})(\d{5})\-(\d{4})$");
            Match matches = rg.Match(nr);
            string group = matches.Groups[1].Value;
            string controlDigits = matches.Groups[3].Value;
            string allDigits = group + matches.Groups[2].Value + controlDigits;
            string nn = "";
            int checkSum = 0;

            if (!matches.Success)
                return (false, null);


            if (Int32.Parse(allDigits.Substring(2, 1)) < 2)
                return (false, null);


            for (int n = 0; n < allDigits.Length; n++)
            {
                nn += ((((n + 1) % 2) + 1) * Int32.Parse(allDigits.Substring(n, 1)));
            }


            for (int n = 0; n < nn.Length; n++)
            {
                checkSum += Int32.Parse(nn.Substring(n, 1));
            }

            return checkSum % 10 == 0 ? (true, matches.Value.ToString()) : (false, null);
        }
    }
}
