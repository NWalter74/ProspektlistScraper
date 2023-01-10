using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace SitemapGenerator.Sitemap
{
    public class Loader : ILoader
    {
        private static HttpClient _client;
        public Loader()
        {
            try
            {
                _client = new HttpClient();
                _client.Timeout = TimeSpan.FromSeconds(300);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }
        public async Task<string> Get(string url)
        {
            try
            {
                HttpResponseMessage resp = await _client.GetAsync(url);
                return await resp.Content.ReadAsStringAsync();
            }
            catch(Exception ex)
            {
                //Fel som tex 443 blir fångade med utan att det castas ett fel så att appen bara fortsätter skrapningen
                //System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            return "";
        }
    }
}
