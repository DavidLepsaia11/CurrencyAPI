using SAPbouiCOM;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using CurrencyTypes.Services;

namespace CurrencyTypes.Controllers
{
    public class CurrencyFormController : BaseFormController
    {
        #region Properties
        static private SBObob bridge { get; set; } 
        public Grid grid { get; set; }
        public CurrencyService currencyService { get; set; }
        public List <string> code { get; set; }
        public List <int> quantity { get; set; }
        public List<string> rateFormated { get; set; }
        public List<string> diffFormated { get; set; }
        public List<double> rate { get; set; }
        public List<string> name { get; set; }
        public List<double> diff { get; set; }
        public List<DateTime> date { get; set; }
        public List<DateTime> validFromDate { get; set; }
        #endregion

        #region Constuctor
        public CurrencyFormController(SAPbobsCOM.Company Company, IForm Form) : base(Company, Form) 
        {      
            grid = ((Grid)(Form.Items.Item("Item_0").Specific));
            bridge = (SBObob)Company.GetBusinessObject(BoObjectTypes.BoBridge);
            
            currencyService = new CurrencyService();
            code = new List<string>();
            quantity = new List<int>();
            rateFormated = new List<string>();
            diffFormated = new List<string>();
            rate = new List<double>();
            name = new List<string>();
            diff = new List<double>();
            date = new List<DateTime>();
            validFromDate = new List<DateTime>();
        }
        #endregion

        #region Methods

        public void GridSetup()
        {

            grid.DataTable.Columns.Add("Code", BoFieldsType.ft_AlphaNumeric);

            grid.DataTable.Columns.Add("Quantity", BoFieldsType.ft_Integer);

            grid.DataTable.Columns.Add("RateFormated", BoFieldsType.ft_AlphaNumeric);

            grid.DataTable.Columns.Add("DiffFormated", BoFieldsType.ft_AlphaNumeric);

            grid.DataTable.Columns.Add("Rate", BoFieldsType.ft_Float);

            grid.DataTable.Columns.Add("Name", BoFieldsType.ft_AlphaNumeric);

            grid.DataTable.Columns.Add("Diff", BoFieldsType.ft_Float);

            grid.DataTable.Columns.Add("Date", BoFieldsType.ft_Date);

            grid.DataTable.Columns.Add("ValidFromDate", BoFieldsType.ft_Date);

            grid.AutoResizeColumns();

        }
   
        public void FillGrid()
        {
            FromAPIToList();
            var pBar = new RSM.Core.SDK.UI.ProgressBar.ProgressBarManager(SAPbouiCOM.Framework.Application.SBO_Application, "resetting", code.Count);

            try
            {
                int startPos = 0;
                pBar.SetValue(startPos);

                Form.Freeze(true);
                for (int i = 0; i < code.Count; i++)
                {
                    grid.DataTable.Rows.Add(1);
                    grid.DataTable.Columns.Item("Code").Cells.Item(i).Value = code[i];
                    grid.DataTable.Columns.Item("Quantity").Cells.Item(i).Value = quantity[i];
                    grid.DataTable.Columns.Item("RateFormated").Cells.Item(i).Value = rateFormated[i];
                    grid.DataTable.Columns.Item("DiffFormated").Cells.Item(i).Value = diffFormated[i];
                    grid.DataTable.Columns.Item("Rate").Cells.Item(i).Value = rate[i];
                    grid.DataTable.Columns.Item("Name").Cells.Item(i).Value = name[i];
                    grid.DataTable.Columns.Item("Diff").Cells.Item(i).Value = diff[i];
                    grid.DataTable.Columns.Item("Date").Cells.Item(i).Value = date[i];
                    grid.DataTable.Columns.Item("ValidFromDate").Cells.Item(i).Value = validFromDate[i];
                    pBar.SetValue(startPos+=1);
                }
            }
            catch (Exception)
            {
                pBar.Stop();
                pBar.Dispose();
            }
            Form.Freeze(false);
            pBar.Stop();
            pBar.Dispose();
        }
    
        private void FromAPIToList()
        {
            foreach (var currency in currencyService.GetCurrencies().currencies)
            {
                code.Add(currency.code);
                quantity.Add(currency.quantity);
                rateFormated.Add(currency.rateFormated);
                diffFormated.Add(currency.diffFormated);
                rate.Add(currency.rate);
                name.Add(currency.name);
                diff.Add(currency.diff);
                date.Add(currency.date);
                validFromDate.Add(currency.validFromDate);
            }
        }
      
        public void FillRate()
        { 
           var selectedCurrency = FindCurrency();
            DateTime Today = DateTime.Today.AddDays(1);
            int Hour = DateTime.Now.Hour;

            try
            {
                if (Hour >= 15)
                {
                foreach (var currency in selectedCurrency)
                {                               
                    bridge.SetCurrencyRate(currency.Key, Today, Convert.ToDouble(currency.Value));                 
                }
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("ვალუტის კურსი განახლდა");
                }
                else
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("17:00 საათამდე არ განახლდება კურსი");
                }        
            }
            catch (Exception e)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("კურსი უკვე  განახლებულია ");
            }

        }
        private IEnumerable<KeyValuePair<string, string>> FindCurrency()
        {
            var selectedCurrency = new Dictionary<string,string>();
            foreach (var currency in currencyService.GetCurrencies().currencies)
            {
                if (currency.code =="USD" || currency.code =="EUR" /*|| currency.code == "GBP"*/)
                {
                    selectedCurrency.Add(currency.code, currency.rateFormated);
                }
            }
            return selectedCurrency;
        }
        #endregion
    }
}
