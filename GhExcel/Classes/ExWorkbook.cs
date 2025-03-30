using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = Microsoft.Office.Interop.Excel;

namespace GhExcel
{
    public class ExWorkbook
    {

        #region members

        public XL.Workbook ComObj = null;

        #endregion

        #region constructors

        public ExWorkbook()
        {
        }

        public ExWorkbook(XL.Workbook comObj)
        {
            this.ComObj = comObj;
        }

        public ExWorkbook(ExWorkbook workbook)
        {
            this.ComObj = workbook.ComObj;
        }

        #endregion

        #region properties

        public virtual ExApp Application
        {
            get { return new ExApp(this.ComObj.Application); }
        }

        public virtual string Name
        {
            get { return System.IO.Path.GetFileNameWithoutExtension(this.ComObj.Name); }
        }

        #endregion

        #region methods

        #region -worksheets

        public ExWorksheet GetWorksheetByName(string name)
        {

            foreach (XL.Worksheet worksheet in this.ComObj.Worksheets)
            {
                if (worksheet.Name == name)
                {
                    return new ExWorksheet(worksheet);
                }
            }

            XL.Worksheet worksheet1 = this.ComObj.Worksheets.Add();
            worksheet1.Name = name;

            return new ExWorksheet(worksheet1); ;
        }

        public ExWorksheet GetActiveWorksheet()
        {
            if (this.ComObj.Worksheets.Count < 1)
            {
                return new ExWorksheet(this.ComObj.Worksheets.Add());
            }
            else
            {
                return new ExWorksheet(this.ComObj.ActiveSheet);
            }
        }

        public List<ExWorksheet> GetAllWorksheets()
        {
            List<ExWorksheet> worksheets = new List<ExWorksheet>();

            foreach (XL.Worksheet sheet in this.ComObj.Worksheets)
            {
                worksheets.Add(new ExWorksheet(sheet));
            }

            return worksheets;
        }

        #endregion

        #endregion

        #region overrides

        public override string ToString()
        {
            return "XL | Workbook {"+this.Name+"}";
        }

        #endregion

    }
}
