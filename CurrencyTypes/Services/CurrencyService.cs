using CurrencyTypes.Model;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace CurrencyTypes.Services
{
   public class CurrencyService
    {
        public Currencies GetCurrencies()
        {
            WebRequest webRequest = WebRequest.Create("https://nbg.gov.ge/gw/api/ct/monetarypolicy/currencies/ka/json");
            HttpWebResponse response = (HttpWebResponse)webRequest.GetResponse();

            if (response.StatusDescription == "OK")
            {
                Stream dataStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                string responseFromServer = reader.ReadToEnd();
                return JsonConvert.DeserializeObject<Currencies[]>(responseFromServer).FirstOrDefault();
            }
            throw new Exception(response.StatusDescription);
        }
    }
}
