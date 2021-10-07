using SAPbouiCOM;
using SAPbobsCOM;
using System;
using System.Linq;
using System.Collections.Generic;
using CurrencyTypes.Services;
using CurrencyTypes.Model;
using System.Drawing;

namespace CurrencyTypes.Services
{
    public class CurrencySetupService
    {

        private SBObob bridge { get; set; }
        public CurrencyService currencyService { get; set; }
        public Model.Currencies Currency { get; set; }
        private Recordset rs { get; set; }

        public CurrencySetupService()
        {
            bridge = (SBObob)RSM.Core.SDK.DI.DIApplication.Company.GetBusinessObject(BoObjectTypes.BoBridge);
            rs = (Recordset)RSM.Core.SDK.DI.DIApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            currencyService = new CurrencyService();
        }
        public void FillRate()
        {
            var selectedCurrency = FindCurrency();

            try
            {
                foreach (var currency in selectedCurrency)
                {
                    if (currency.validFromDate.DayOfWeek == DayOfWeek.Friday)
                    {
                        while (currency.validFromDate.DayOfWeek != DayOfWeek.Monday)
                        {
                            currency.validFromDate = currency.validFromDate.AddDays(1);
                            bridge.SetCurrencyRate(currency.code, currency.validFromDate, Convert.ToDouble(currency.rateFormated), true);
                        }
                    }
                    else
                    {
                        bridge.SetCurrencyRate(currency.code, currency.validFromDate, Convert.ToDouble(currency.rateFormated), true);
                    }
                }
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("ვალუტის კურსი განახლდა");
            }
            catch (Exception e)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(e.Message);
            }

        }

        private IEnumerable<Currency> FindCurrency()
        {
            var CurrencyCodes = currencyService.GetCurrencies().currencies;
            var ISOCurrency = GetCurrencyCode();
            var selectedCurrency = new List<Currency>();
            for (int i = 0; i < CurrencyCodes.Count; i++)
            {
                if (ISOCurrency.Where(x => x == CurrencyCodes[i].code).Count() != 0)
                {
                    selectedCurrency.Add(CurrencyCodes[i]);
                }
            }
            return selectedCurrency;
        }

        private List<string> GetCurrencyCode()
        {
            var tempList = new List<string>();
            rs.DoQuery("SELECT ISOCurrCod FROM OCRN WHERE ISOCurrCod != 'GEL'");
            while (!rs.EoF)
            {
                tempList.Add(rs.Fields.Item("ISOCurrCod").Value.ToString());
                rs.MoveNext();
            }
            return tempList;
        }
    }
}
