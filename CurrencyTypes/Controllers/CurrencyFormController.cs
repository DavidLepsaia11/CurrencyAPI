using SAPbouiCOM;
using SAPbobsCOM;
using System;
using System.Linq;
using System.Collections.Generic;
using CurrencyTypes.Services;
using CurrencyTypes.Model;
using System.Drawing;

namespace CurrencyTypes.Controllers
{
    public class CurrencyFormController : BaseFormController
    {
        #region Properties
        static private SBObob bridge { get; set; } 
        public Grid grid { get; set; }
        public CurrencyService currencyService { get; set; }
        public Model.Currencies Currency { get; set; }
      
        private Recordset rs { get; set; }
        #endregion

        #region Constuctor
        public CurrencyFormController(SAPbobsCOM.Company Company, IForm Form) : base(Company, Form) 
        {      
            grid = ((Grid)(Form.Items.Item("Item_0").Specific));

            bridge = (SBObob)Company.GetBusinessObject(BoObjectTypes.BoBridge);
            
            currencyService = new CurrencyService();      
            rs = (Recordset)Company.GetBusinessObject(BoObjectTypes.BoRecordset);
        }
        #endregion

        #region Methods

        public void GridSetup()
        {

            grid.DataTable.Columns.Add("Code", BoFieldsType.ft_AlphaNumeric);

            grid.DataTable.Columns.Add("Quantity", BoFieldsType.ft_Integer);

            grid.DataTable.Columns.Add("RateFormated", BoFieldsType.ft_AlphaNumeric);

            grid.DataTable.Columns.Add("DiffFormated", BoFieldsType.ft_AlphaNumeric);

            grid.DataTable.Columns.Add("Rate", BoFieldsType.ft_Rate);

            grid.DataTable.Columns.Add("Name", BoFieldsType.ft_AlphaNumeric);

            grid.DataTable.Columns.Add("Diff", BoFieldsType.ft_Rate);

            grid.DataTable.Columns.Add("Date", BoFieldsType.ft_Date);

            grid.DataTable.Columns.Add("ValidFromDate", BoFieldsType.ft_Date);

            grid.DataTable.Columns.Add("SAP Currency", BoFieldsType.ft_Rate);

            grid.AutoResizeColumns();

        }  
        public void FillGrid()
        {
            FromAPIToList();
            var sapCurrencies =  DoQuearyFromRecordSet();
            var pBar = new RSM.Core.SDK.UI.ProgressBar.ProgressBarManager(SAPbouiCOM.Framework.Application.SBO_Application, "resetting", Currency.currencies.Count);
            var CurrencyCode = GetCurrencyCode();
            try
            {
                int startPos = 0;
                pBar.SetValue(startPos);
                Form.Freeze(true);
            
                for (int i = 0; i < Currency.currencies.Count; i++)
                {
                    grid.DataTable.Rows.Add(1);
                    grid.DataTable.Columns.Item("Code").Cells.Item(i).Value = Currency.currencies[i].code;               
                    grid.DataTable.Columns.Item("Quantity").Cells.Item(i).Value = 1;
                    grid.DataTable.Columns.Item("Rate").Cells.Item(i).Value = Currency.currencies[i].rate / Currency.currencies[i].quantity;
                    grid.DataTable.Columns.Item("RateFormated").Cells.Item(i).Value = (Currency.currencies[i].rate / Currency.currencies[i].quantity).ToString();                                               
                    grid.DataTable.Columns.Item("DiffFormated").Cells.Item(i).Value = Currency.currencies[i].diffFormated;
                    grid.DataTable.Columns.Item("Name").Cells.Item(i).Value = Currency.currencies[i].name;
                    grid.DataTable.Columns.Item("Diff").Cells.Item(i).Value = Currency.currencies[i].diff;
                    if (Currency.currencies[i].diff > 0)
                    {
                        int redBackColor = Color.Red.R | (Color.Red.G << 8) | (Color.Red.B << 16);
                        grid.CommonSetting.SetRowBackColor(i + 1, redBackColor);
                    }
                    else if (Currency.currencies[i].diff < 0)
                    {
                        int redBackColor = Color.Green.R | (Color.Green.G << 8) | (Color.Green.B << 16);
                        grid.CommonSetting.SetRowBackColor(i + 1, redBackColor);
                    }
                    grid.DataTable.Columns.Item("Date").Cells.Item(i).Value = Currency.currencies[i].date;
                    grid.DataTable.Columns.Item("ValidFromDate").Cells.Item(i).Value = Currency.currencies[i].validFromDate;
                
                    if(sapCurrencies.ContainsKey(Currency.currencies[i].code))
                    {
                        grid.DataTable.Columns.Item("SAP Currency").Cells.Item(i).Value = sapCurrencies[Currency.currencies[i].code];
                    }
                    else
                    {
                        grid.DataTable.Columns.Item("SAP Currency").Cells.Item(i).Value = -1;
                    }
               
                    pBar.SetValue(startPos+=1);
                }
            }
            catch (Exception e)
            {
                pBar.Stop();
                pBar.Dispose();
            }
            Form.Freeze(false);
            pBar.Stop();
            pBar.Dispose();
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
                       bridge.SetCurrencyRate(currency.code , currency.validFromDate, Convert.ToDouble(currency.rateFormated), true);                 
                    }                              
                }
                    SAPbouiCOM.Framework.Application.SBO_Application.MessageBox("ვალუტის კურსი განახლდა");          
            }
            catch (Exception e)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(e.Message);
            }

        }

        #region Private Helper Methods
        private void FromAPIToList()
        {
            Currency = currencyService.GetCurrencies();
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

        private Dictionary<string,double> DoQuearyFromRecordSet()
        {            
            var ValidFromDate = currencyService.GetCurrencies().currencies.Select(x => x.validFromDate).ToList();
            var tempList = new Dictionary<string, double>();
            rs.DoQuery($"SELECT * FROM ORTT where RateDate = '{ValidFromDate[0]}'");
            while (!rs.EoF)
            {
                tempList.Add(rs.Fields.Item("Currency").Value.ToString(), double.Parse(rs.Fields.Item("Rate").Value.ToString()));
                rs.MoveNext();
            }
            return tempList;
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
        
        #endregion

        #endregion
    }
}
