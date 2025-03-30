using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using XL = Microsoft.Office.Interop.Excel;

namespace GhExcel
{
    public class ExWorksheet
    {

        #region members

        public XL.Worksheet ComObj = null;

        #endregion

        #region constructors

        public ExWorksheet()
        {
        }

        public ExWorksheet(XL.Worksheet comObj)
        {
            this.ComObj = comObj;
        }

        public ExWorksheet(ExWorksheet worksheet)
        {
            this.ComObj = worksheet.ComObj;
        }

        #endregion

        #region properties

        public virtual ExWorkbook Workbook
        {
            get { return new ExWorkbook((XL.Workbook)(this.ComObj.Parent)); }
        }

        public virtual ExApp Application
        {
            get { return new ExApp(this.ComObj.Application); }
        }

        #endregion

        #region methods

        #region -ranges

        public ExRange GetRange(string minAddress, string maxAddress)
        {
            return new ExRange(this.ComObj.Range[minAddress, maxAddress]);
        }

        public ExRange WriteData(List<List<GH_String>> data, string address)
        {
            int y = data[0].Count;
            int x = data.Count;

            string[,] values = new string[y, x];
            double[,] numbers = new double[y, x];

            bool isNumeric = true;
            bool isFunction = true;

            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    values[j, i] = data[i][j].Value;
                    if (double.TryParse(values[j, i], out double num)) numbers[j, i] = num; else isNumeric = false;
                    if (!(values[j, i].ToCharArray()[0].Equals('='))) isFunction = false;
                }
            }

            string max = address.Move(x-1,y-1);
            ExRange range =this.GetRange(address, max);

            if (isNumeric)
            {
                range.SetNumericValues(numbers);
            }
            else if (isFunction)
            {
                range.SetFormula(values);
            }
            else
            {
                range.SetTextValues(values);
            }

            return range;
        }

        #endregion

        #endregion

        #region overrides

        public override string ToString()
        {
            return "XL | Worksheet {" + this.ComObj.Name + "}";
        }

        #endregion
    }
}
