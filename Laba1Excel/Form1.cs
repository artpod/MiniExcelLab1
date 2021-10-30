using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Laba1Excel
{

    public partial class Form1 : Form
    {

        const int st = 100;
        int COLUMNS = 0;
        int ROWS = 0;
        private string _currentFilePath = "";
        CreateNewEl NewEL = new CreateNewEl();
        Dictionary<int, object> _values = new Dictionary<int, object>();
        Cell[,] table = new Cell[st, st];
        Parser1 parser = new Parser1();
        public Form1()
        {
            InitializeComponent();
            for (int i = 0; i < st; i++)
                for (int j = 0; j < st; j++)
                    table[i, j] = new Cell();

        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            COLUMNS = 5;
            ROWS = 10;
            string temp = null;
            char c = 'A';
            Cell[,] table = new Cell[st, st];
            for (int i = 0; i < COLUMNS; i++)
            {
                temp += c;
                dataGridView1.Columns.Add(Name, temp);
                c++;
                temp = null;
            }

            for (int i = 0; i < ROWS; i++)
                dataGridView1.Rows.Add();

            for (int i = 0; i < (dataGridView1.Rows.Count); i++)
            {
                dataGridView1.Rows[i].HeaderCell.Value = i.ToString();
            }
            for (int i = 0; i < (dataGridView1.Rows.Count); i++)
                for (int j = 0; j < (dataGridView1.Columns.Count); j++)
                    dataGridView1.Rows[i].Cells[j].Value = 0;

        }

        /* private void SaveDataGridView(string filePath)
         {
             _currentFilePath = filePath;
             dataGridView1.EndEdit();

             DataTable table = new DataTable("data");
             ForgeDataTable(table);
             table.WriteXml(filePath);
         }

         private void ForgeDataTable(DataTable table)
         {
             foreach (DataGridViewColumn dgvColumn in dataGridView1.Columns)
             {
                 table.Columns.Add(dgvColumn.Index.ToString());
             }

             foreach(DataGridViewRow dgvRow in dataGridView1.Rows)
             {
                 DataRow dataRow = table.NewRow();
                 foreach (DataColumn col in table.Columns)
                 {
                     Cell cell = (Cell)dgvRow.Cells[Int32.Parse(col.ColumnName)].Tag;
                     dataRow[col.ColumnName] = Convert.ToDouble(cell.value_1);
                 }
                 table.Rows.Add(dataRow);
             }
         }

         private bool SaveDataGridView(string filePath, string dummy)
         {
             if (!filePath.Equals(""))
             {
                 SaveDataGridView(filePath);
                 return true;
             }
             else if (saveFileDialog.ShowDialog() == DialogResult.OK)
             {
                 SaveDataGridView(saveFileDialog.FileName);
                 return true;
             }
             return false;
         }

         private void LoadDataGridView(string filePath)
         {
             _currentFilePath = filePath;
             DataSet dataSet = new DataSet();
             dataSet.ReadXml(filePath);
             DataTable table = dataSet.Tables[0];

             dataGridView1.ColumnCount = table.Columns.Count;
             dataGridView1.RowCount = table.Rows.Count;

             foreach(DataGridViewRow dgvRow in dataGridView1.Rows)
             {
                 foreach(DataGridViewCell dgvCell in dgvRow.Cells)
                 {
                     string cellName = "R" + (dgvRow.Index + 1).ToString() + "C" + (dgvCell.ColumnIndex + 1).ToString();
                     string formula = table.Rows[dgvCell.RowIndex][dgvCell.ColumnIndex].ToString();
                     dgvCell.Tag = new Cell();
                 }

             }

         }
        */

        static int CalcKey(int rowIndex, int columnIndex)
        {
            return rowIndex + (columnIndex << 16);
        }

        private void AddColumn_Click(object sender, EventArgs e)
        {
            NewEL.AddColumn(dataGridView1, COLUMNS);
            COLUMNS++;
            for (int i = 0; i < (dataGridView1.Rows.Count); i++)
                dataGridView1.Rows[i].Cells[COLUMNS - 1].Value = 0;
        }

        private void AddRow_Click(object sender, EventArgs e)
        {
            NewEL.AddRow(ROWS, dataGridView1);
            ROWS++;
            for (int i = 0; i < (dataGridView1.Columns.Count); i++)
                dataGridView1.Rows[ROWS - 1].Cells[i].Value = 0;
        }
        //from this
        private void DeleteRow_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount <= 2) { MessageBox.Show("Подальше видалення неможливе"); return; }
            for (int i = 0; i < COLUMNS; i++)
            {
                addNewValue("0", ROWS - 1, i);
            }
            dataGridView1.Rows.RemoveAt(ROWS - 1);
            ROWS--;
            dataGridView1.Rows[ROWS].HeaderCell.Value = ROWS.ToString();
        }

        private void DeleteColumn_Click(object sender, EventArgs e)
        {
            bool columnIsEmpty = true;
            /* for (int j = 0; j < dataGridView1.RowCount; j++)
             {
                 if (dataGridView1.Rows[j].Cells[COLUMNS - 1].Value != null)
                 {
                     columnIsEmpty = false;
                 }
             }
             if (columnIsEmpty == false)
             {
                 MessageBox.Show("Подальше видалення неможливе. Колонка заповнена");
             }
             if (columnIsEmpty == true)
             {*/
            if (dataGridView1.ColumnCount <= 2) { MessageBox.Show("Подальше видалення неможливе"); return; }
            for (int i = 0; i < ROWS; i++)
                {
                    addNewValue("0", i, COLUMNS - 1);
                }
                dataGridView1.Columns.RemoveAt(COLUMNS - 1);
                COLUMNS--;
           // }
            //if (dataGridView1.ColumnCount <= 2) { MessageBox.Show("Подальше видалення неможливе"); return; }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = table[dataGridView1.CurrentCell.RowIndex, dataGridView1.CurrentCell.ColumnIndex].exp;
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            string temp;
            if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
            {
                temp = Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
            }
            else temp = "0";
            int r = dataGridView1.CurrentCell.RowIndex;
            int c = dataGridView1.CurrentCell.ColumnIndex;
            if (temp != null) addNewValue(temp, r, c);
        }
        

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string temp = (string)textBox1.Text;
                int r = dataGridView1.CurrentCell.RowIndex;
                int c = dataGridView1.CurrentCell.ColumnIndex;
                addNewValue(temp, r, c);
            }
        }

        // up to this
        public void addNewValue(string exptession, int r, int c)
        {
            table[r, c].exp = exptession;
            int h = 0;
            if (exptession != null)
            {
                while (h < exptession.Length)
                {
                    string str = null;
                    int t2, t1 = (int)exptession[h];
                    str += exptession[h];
                    h++;
                    if ((t1 > 64) && (t1 < 91))
                    {
                        str += exptession[h];
                        t2 = (int)exptession[h];
                        h++;
                        table[t2 - 48, t1 - 65].dependents.Add(table[r, c].getName(c, r));
                        exptession = exptession.Replace(str, table[t2 - 48, t1 - 65].value_1);
                    }
                }
                if (circle(table[r, c], r, c))
                {
                    // Result Res = parser.Evaluate(exptession);
                    try
                    {
                        double Res = parser.Evaluate(exptession);

                        //  if (Res)
                        // {
                        table[r, c].value_1 = Res.ToString();
                        update(table[r, c], r, c);
                        dataGridView1.Rows[r].Cells[c].Value = Res.ToString();
                        // }
                        // else dataGridView1.Rows[r].Cells[c].Value = Res.GetValue();
                    }
                    catch
                    {
                        MessageBox.Show("Сталася помилка. Спробуйте ще раз!");
                    }
                }
            }
        }
        public void update(Cell cell, int r, int c)
        {
            for (int i = 0; i < cell.dependents.Count; i++)
            {
                int t2 = (int)cell.dependents[i][1] - 48, t1 = (int)cell.dependents[i][0] - 65;
                Conection(table[t2, t1].exp, t2, t1);

            }
        }
        public bool circle(Cell cell, int r, int c)
        {
            for (int i = 0; i < cell.dependents.Count; i++)
            {
                int t2 = (int)cell.dependents[i][1] - 48, t1 = (int)cell.dependents[i][0] - 65;
                for (int j = 0; j < table[t2, t1].dependents.Count; j++)
                    if (table[t2, t1].dependents[j] == table[r, c].getName(r, c)) {
                        nameCircle(cell, r, c);
                        nameCircle(table[t2, t1], t2, t1);
                        table[t2, t1].dependents.RemoveAt(j);
                        return false;
                    }
            }
            return true;
        }
        public void nameCircle(Cell cell, int r, int c)
        {
            for (int fl = 0; fl < cell.dependents.Count; fl++) {
                int _t2 = (int)cell.dependents[fl][1] - 48, _t1 = (int)cell.dependents[fl][0] - 65;
                dataGridView1.Rows[_t2].Cells[_t1].Value = "#CIRCLE";
            }
        }

        public void Conection(string temp, int row, int colum)
        {
            int h1 = 0;
            while (h1 < temp.Length)
            {
                string str = null;
                int t2, t1 = (int)temp[h1];
                str += temp[h1];
                h1++;
                if (t1 > 61 && t1 < 91)
                {
                    str += temp[h1];
                    t2 = (int)temp[h1];
                    h1++;
                    temp = temp.Replace(str, table[t2 - 48, t1 - 65].value_1);
                }
            }
            if (circle(table[row, colum], row, colum))
            {
                // Result Res = parser.Evaluate(temp);
                try
                {
                    double Res = parser.Evaluate(temp);
                    // if (Res.except())
                    //{
                    table[row, colum].value_1 = Res.ToString();
                    update(table[row, colum], row, colum);
                    dataGridView1.Rows[row].Cells[colum].Value = Res.ToString();
                }
                catch
                {
                    MessageBox.Show("Сталася помилка. Спробуйте ще раз!");
                }
                // }
                // else return;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Ви впевнені? Що бажаєте закрити файл?", "Повідомлення", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation))
            {
                //Application.Exit();
            }
            else
                e.Cancel = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Ця програма виконує наступні бінарні операції: +,-,*,/, возведення в степінь ^ та обчислення max,min,mod,div");
        }

        private void OpenStripMenuItem_Click(object sender, EventArgs e)
        {
            Stream mystr = null;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                if ((mystr = openFileDialog.OpenFile()) != null)
                {
                    using (mystr)
                    {
                        try
                        {
                            StreamReader sr = new StreamReader(mystr);
                            string scr = sr.ReadLine();
                            string scc = sr.ReadLine();
                            int cr = Convert.ToInt32(scr);
                            int cc = Convert.ToInt32(scc);
                            for (int i = 0; i < cr; i++)
                                for (int j = 0; j < cc; j++)
                                {
                                    string cell_name = (char)(j + 65) + (i + 1).ToString();
                                    dataGridView1.Rows[i].Cells[j].Value = null;
                                    table[i, j].value_1 = null;
                                    //RefreshCells();
                                    dataGridView1.Rows[i].Cells[j].Value = sr.ReadLine();
                                    table[i, j].value_1 = (string)dataGridView1.Rows[i].Cells[j].Value;
                                   // table[dataGridView1.CurrentCell.RowIndex, dataGridView1.CurrentCell.ColumnIndex].exp = 
                                }
                            for (int i = 0; i < cr; i++)
                                for (int j = 0; j < cc; j++)
                                {
                                    string cell_name = (char)(j + 65) + (i + 1).ToString();
                                    table[i, j].exp = sr.ReadLine();
                                   // string temp = (string)dataGridView1.Rows[i].Cells[j].Value;
                                    //int r = dataGridView1.CurrentCell.RowIndex;
                                   // int c = dataGridView1.CurrentCell.ColumnIndex;
                                    //if (temp != null) addNewValue(temp, r, c);
                                    
                                }
                            sr.Close();
                        }
                        catch { }
                    }
                }
            }
        }

        private void SaveStripMenuItem_Click(object sender, EventArgs e)
        {
            Stream mystream;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if ((mystream = saveFileDialog.OpenFile()) != null)
                {
                    StreamWriter sw = new StreamWriter(mystream);
                    sw.WriteLine(dataGridView1.RowCount);
                    sw.WriteLine(dataGridView1.ColumnCount);
                    for (int i=0; i < dataGridView1.RowCount; i++)
                    {
                        for (int j = 0; j < dataGridView1.ColumnCount; j++){
                            if (dataGridView1.Rows[i].Cells[j].Value != null)
                                sw.WriteLine(dataGridView1.Rows[i].Cells[j].Value.ToString());
                            else sw.WriteLine("");
                        }
                    }
                    try
                    {
                        for (int i = 0; i < dataGridView1.RowCount; i++)
                        {
                            for (int j = 0; j < dataGridView1.ColumnCount; j++)
                            {
                               // string cell_name = (char)(j + 65) + (i + 1).ToString();
                                if (table[i, j].value_1 != null)
                                    sw.WriteLine(table[i, j].value_1);
                                else sw.WriteLine("");
                            }
                        }
                    }
                    catch { }
                    sw.Close();
                    mystream.Close();
                }
            }
        }
    }


}

