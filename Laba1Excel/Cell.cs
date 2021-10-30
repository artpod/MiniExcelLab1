using System;
using System.Collections.Generic;
using System.Text;

namespace Laba1Excel
{
    public class Cell
    {
        public string value_1;
        public string exp;
        public List<string> dependents = new List<string>();
        public Cell()
        {
            value_1 = "0";
        }
        public string getName(int column, int row)
        {
            string temp = null;
            temp += (char)(column + 65);
            temp += (char)(row + 48);
            return temp;
        }
    }
}
