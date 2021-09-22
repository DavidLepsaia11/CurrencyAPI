using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;
using CurrencyTypes.Controllers;

namespace CurrencyTypes
{
    [FormAttribute("CurrencyTypes.CurrencyTypes", "Forms/CurrencyTypes.b1f")]
    class CurrencyTypes : UserFormBase
    {
        public CurrencyTypes()
        {
        }
        CurrencyFormController currencyController { get; set; }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
            bool Iseven_Button0_PressedAfterExecute = false;
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_0").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_3").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            
        }

        private SAPbouiCOM.Grid Grid0;

        private void OnCustomInitialize()
        {
            currencyController = new CurrencyFormController(RSM.Core.SDK.DI.DIApplication.Company, UIAPIRawForm);
            currencyController.GridSetup();
        }

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (Iseven_Button0_PressedAfterExecute != true)
            {
               currencyController.FillGrid();
                Iseven_Button0_PressedAfterExecute = true;
            }
        }

        private SAPbouiCOM.Button Button2;

        private void Button2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            currencyController.FillRate();
        }
    }
}