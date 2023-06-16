using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtensions
{
    class Xl
    {
        public static int GetCellRowIndex(string cellName)
        {
            int rowIndex = 0;
            int i = cellName.Length - 1;
            int multiplier = 1;
            while (i >= 0 && cellName[i] >= '0' && cellName[i] <= '9')
            {
                rowIndex += (cellName[i] - '0') * multiplier;
                multiplier *= 10;
                i--;
            }
            return rowIndex;
        }

        public static int GetCellColumnIndex(string cellName) 
        {
            const int lettersInAlphabet = 26;
            const int firstLetterMinusOne = 'A' - 1;

            int columnIndex = 0;
            int multiplier = 1;
            int i = cellName.Length - 1;
            while (i >= 0 && cellName[i] >= '0' && cellName[i] <= '9') i--;
            while (i >= 0
                && (cellName[i] >= 'a' && cellName[i] <= 'z' || cellName[i] >= 'A' && cellName[i] <= 'Z'))
            {
                columnIndex += (char.ToUpper(cellName[i]) - firstLetterMinusOne) * multiplier;
                multiplier *= lettersInAlphabet;
                i--;
            }
            return columnIndex;
        }

        public static string GetCellColumnLetter(string cellName)
        {
            var columnLetter = new StringBuilder();
            int i = 0;
            while (i < cellName.Length
                && (cellName[i] >= 'a' && cellName[i] <= 'z' || cellName[i] >= 'A' && cellName[i] <= 'Z'))
                columnLetter.Append(char.ToUpper(cellName[i++]));
            return columnLetter.ToString();
        }
    }
}

