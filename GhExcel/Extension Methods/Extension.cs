using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GhExcel
{
    public static class Extension
    {

        public static string[] Flatten(this string[,] input)
        {
            List<string> output = new List<string>();

            for (int i = 0; i < input.GetLength(0); i++)
            {
                for (int j = 0; j < input.GetLength(1); j++)
                {
                    output.Add(input[i, j]);
                }
            }

            return output.ToArray();
        }

        public static string ToAddress(this int[] input)
        {
            int col = input[0];
            string colLetter = String.Empty;
            int mod;

            while (col > 0)
            {
                mod = (col - 1) % 26;
                colLetter = ((char)(65 + mod)).ToString() + colLetter;
                col = (int)((col - mod) / 26);
            }

            string address = (colLetter + input[1]);
            return address;
        }

        public static int[] ToLocation(this string input)
        {

            char[] arrC = input.ToCharArray();
            string Lint = "";
            string Lstr = "";
            int retVal = 0;

            for (int i = 0; i < arrC.Length; i++)
            {
                if (char.IsNumber(arrC[i]))
                {
                    Lint += arrC[i];
                }
                else
                {
                    Lstr += arrC[i];
                }
            }

            string col = Lstr.ToUpper();
            int k = col.Length - 1;

            for (int i = 0; i < k + 1; i++)
            {
                char colPiece = col[i];
                int t = (int)colPiece;
                int colNum = t - 64;
                retVal = retVal + colNum * (int)(Math.Pow(26, col.Length - (i + 1)));
            }

            return new int[] { retVal, Convert.ToInt32(Lint) };
        }

        public static string Move(this string input, int col, int row)
        {
            int[] source = input.ToLocation();
            int[] temp = { source[0] +col,source[1]+row};

            return temp.ToAddress();
        }

    }
}
