/*using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace Laba1Excel
{
    class Result
    {
        //cout errors
        public Double Value;
        public Parser.Errors Code;

        public Result()
        {
            Value = 0.0;
            Code = Parser.Errors.NOERR;
        }
        public Result(double V, Parser.Errors code)
        {
            Value = V;
            Code = code;
        }
        public bool except()
        {
            return (Code == Parser.Errors.NOERR);
        }
        public string GetValue()
        {
            switch (Code)
            {
                case Parser.Errors.NOERR: return Value.ToString(); ;
                case Parser.Errors.DIVBYZERO: { MessageBox.Show("Ділення на нуль неможливе. Змініть значення"); return "#ERROR"; }
                case Parser.Errors.NOEXP: return "#ERROR";
                case Parser.Errors.UNBALPARENS: return "#ERROR";
                case Parser.Errors.SYNTAX: return "#ERROR";
            }
            return "";
        }
    }

}
*/