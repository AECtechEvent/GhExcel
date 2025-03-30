using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = Microsoft.Office.Interop.Excel;

namespace GhExcel
{
    public class ExRange
    {

        #region members

        public XL.Range ComObj = null;

        #endregion

        #region constructors

        public ExRange()
        {
        }

        public ExRange(XL.Range comObj)
        {
            this.ComObj = comObj;
        }

        public ExRange(ExRange range)
        {
            this.ComObj = range.ComObj;
        }

        #endregion

        #region properties

        public virtual ExWorksheet Worksheet
        {
            get { return new ExWorksheet(this.ComObj.Worksheet); }
        }

        public virtual ExWorkbook Workbook
        {
            get { return new ExWorkbook(this.Worksheet.Workbook); }
        }

        public virtual ExApp Application
        {
            get { return new ExApp(this.ComObj.Application); }
        }

        public virtual int MinColumn
        {
            get { return this.ComObj.Column; }
        }

        public virtual int MaxColumn
        {
            get { return this.ComObj.Columns[this.ComObj.Columns.Count].Column; }
        }

        public virtual int MinRow
        {
            get { return this.ComObj.Row; }
        }

        public virtual int MaxRow
        {
            get { return this.ComObj.Rows[this.ComObj.Rows.Count].Row; }
        }

        public virtual string Min
        {
            get
            {
                int[] location = { this.MinColumn, this.MinRow };
                return location.ToAddress();
            }
        }

        public virtual string Max
        {
            get
            {
                int[] location = { this.MaxColumn, this.MaxRow };
                return location.ToAddress();
            }
        }

        protected int[] ExtentArray
        {
            get { 

                return new int[] { this.MinColumn, this.MinRow, this.MaxColumn, this.MaxRow }; 
            }
        }

        #endregion

        #region methods

        public void SetFormula(string[,] values)
        {
            string[] formulas = values.Flatten();
            int i = 0;
            foreach (XL.Range cell in this.ComObj.Cells)
            {
                cell.Formula = formulas[i++];
            }
        }

        public void SetTextValues(string[,] values)
        {
            this.ComObj.NumberFormat = "@";
            this.ComObj.Value2 = values;
        }

        public void SetNumericValues(double[,] values)
        {
            this.ComObj.NumberFormat = "0.00";
            this.ComObj.Value2 = values;
        }

        public GH_Structure<GH_String> ReadData(GH_Path path)
        {
            GH_Structure<GH_String> ghData = new GH_Structure<GH_String>();

            int[] L = this.ExtentArray;

            System.Array values = (System.Array)this.ComObj.Value2;

            if (values != null)
            {
                for (int i = 1; i < (L[2] - L[0] + 2); i++)
                {
                    for (int j = 1; j < (L[3] - L[1] + 2); j++)
                    {
                        string val = string.Empty;
                        if (values.GetValue(j, i) != null) val = values.GetValue(j, i).ToString();
                        ghData.Append(new GH_String(val), path.AppendElement(i));
                    }
                }
            }

            return ghData;
        }

        #endregion

        #region overrides

        public override string ToString()
        {
            return "XL | Range{"+ this.Min+":" + this.Max +"}";
        }

        #endregion

    }
}
