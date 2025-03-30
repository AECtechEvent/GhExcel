using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = Microsoft.Office.Interop.Excel;

namespace GhExcel
{
    public class ExApp
    {

        #region members

        public XL.Application ComObj = null;

        #endregion

        #region constructors

        public ExApp()
        {
            try
            {
                this.ComObj = (XL.Application)Marshal2.GetActiveObject("Excel.Application");
            }
            catch (Exception e)
            {
                this.ComObj = new XL.Application();
            }
            if (this.ComObj.ActiveWorkbook == null) this.ComObj.Workbooks.Add();
            if (!this.ComObj.Visible) this.ComObj.Visible = true;
        }

        public ExApp(ExApp exApp)
        {
            this.ComObj = exApp.ComObj;
        }

        public ExApp(XL.Application comObj)
        {
            this.ComObj = comObj;
        }

        #endregion

        #region properties



        #endregion

        #region methods

        #region -workbooks

        public ExWorkbook LoadWorkbook(string filePath)
        {
            ExWorkbook workbook = new ExWorkbook(this.ComObj.Workbooks.Open(filePath));

            return workbook;
        }

        public ExWorkbook GetActiveWorkbook()
        {

            if (this.ComObj.Workbooks.Count < 1)
            {
                //Creates a new workbook if no workbook(s) are open
                return new ExWorkbook(this.ComObj.Workbooks.Add());
            }
            else
            {
                //Gets the topmost workbook if workbook(s) are open
                return new ExWorkbook(this.ComObj.ActiveWorkbook);
            }

        }

        public List<ExWorkbook> GetAllWorkbooks()
        {
            List<ExWorkbook> output = new List<ExWorkbook>();

            foreach (XL.Workbook workbook in this.ComObj.Workbooks)
            {
                output.Add(new ExWorkbook(workbook));
            }

            return output;
        }

        #endregion

        #endregion

        #region overrides

        public override string ToString()
        {
            return "XL | App";
        }

        #endregion

    }
}
