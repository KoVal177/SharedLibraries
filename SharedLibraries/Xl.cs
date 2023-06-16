using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtensions
{
    class Xl
    {
        public static int GetCellRowIndex(string cell)
        {
            int rowIndex = 0;
            int i = cell.Length - 1;
            int multiplier = 1;
            while (i >= 0 && cell[i] >= '0' && cell[i] <= '9')
            {
                rowIndex += (cell[i] - '0') * multiplier;
                multiplier *= 10;
                i--;
            }
            return rowIndex;
        }

        public static int GetCellColumnIndex(string cell) 
        {
            const int lettersInAlphabet = 26;
            const int firstLetterMinusOne = 'A' - 1;

            int columnIndex = 0;
            int multiplier = 1;
            int i = cell.Length - 1;
            while (i >= 0 && cell[i] >= '0' && cell[i] <= '9') i--;
            while (i >= 0
                && (cell[i] >= 'a' && cell[i] <= 'z' || cell[i] >= 'A' && cell[i] <= 'Z'))
            {
                columnIndex += (char.ToUpper(cell[i]) - firstLetterMinusOne) * multiplier;
                multiplier *= lettersInAlphabet;
                i--;
            }
            return columnIndex;
        }

        public static string GetCellColumnLetter(string cell)
        {
            var columnLetter = new StringBuilder();
            int i = 0;
            while (i < cell.Length
                && (cell[i] >= 'a' && cell[i] <= 'z' || cell[i] >= 'A' && cell[i] <= 'Z'))
                columnLetter.Append(char.ToUpper(cell[i++]));
            return columnLetter.ToString();
        }
    }
}

