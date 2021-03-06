using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace Laba1Excel
{
    class CreateNewEl
    {
        const int b = 65;
        const int c = 26;
        char letter = 'A';
        char firstletter = 'A';
        int nfirstletter = 1;
        string temp;
        int r;
        public int AddColumn(DataGridView dgv, int countColums)
        {
            if (countColums < c)
            {
                int t = b + countColums;
                letter = (char)t;
                temp += letter;
            }
            else
            {
                int y = 0;
                while (countColums >= c)
                {
                    countColums -= c;
                    ++y;
                }
                firstletter = (char)(y + (b - 1));
                letter = (char)(countColums + b);
                temp += firstletter;
                temp += letter;
            }
            DataGridViewColumn col = (DataGridViewColumn)dgv.Columns[0].Clone();
            col.HeaderCell.Value = temp;

            dgv.Columns.Add(col);
            temp = null;

            if (firstletter != 'Z')
                if (letter != 'Z')
                    ++letter;
                else
                {
                    letter = 'A';
                    firstletter++;
                }
            else
            {
                firstletter = 'A';
                ++nfirstletter;
            }
            return 0;
        }

        public int AddRow(int _c, DataGridView dgv)
        {
            r = _c;
            ++_c;
            string t = null; t += r;
            DataGridViewRow row = (DataGridViewRow)dgv.Rows[r].Clone();
            dgv.Rows[r].HeaderCell.Value = _c.ToString();
            dgv.Rows.Add(row);
            return 0;
        }
    }
}
