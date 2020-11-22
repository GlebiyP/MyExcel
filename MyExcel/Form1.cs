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

namespace MyExcel
{
    public partial class Form1 : Form
    {
        int currRow, currCol;
        int ROWS = 10;
        int COLUMNS = 10;
        public static Dictionary<string, Cell> table = new Dictionary<string, Cell>();
        Calculator ExcelCalculator = new Calculator(table);
        public Form1()
        {
            InitializeComponent();

            addRowButton.Click += new EventHandler(addRowButton_Click);
            addColumnButton.Click += new EventHandler(addColumnButton_Click);
            deleteRowButton.Click += new EventHandler(deleteRowButton_Click);
            deleteColumnButton.Click += new EventHandler(deleteColumnButton_Click);
            toolStripMenuItemAbout.Click += new EventHandler(toolStripMenuItemAbout_Click);
            saveStripMenuItem.Click += new EventHandler(saveStripMenuItem_Click);
            openStripMenuItem.Click += new EventHandler(openStripMenuItem_Click);
            buttonEnter.Click += new EventHandler(buttonEnter_Click);

            CreateTable(ROWS, COLUMNS);
        }
       
        private void Form1_Load_1(object sender, EventArgs e)
        {
          
        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void addRowButton_Click(object sender, EventArgs e)
        {
            //добавление ряда
            AddRow(dgv);
            RefreshCell();
        }

        private void addColumnButton_Click(object sender, EventArgs e)
        {
            //добавление колонки
            AddColumn(dgv);
            RefreshCell();
        }

        private void deleteRowButton_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Ви впевнені, що хочете видалити рядок? Він може містити важливі дані.",
                "Delete row", MessageBoxButtons.YesNo);
            //удаление ряда
            if (dialogResult == DialogResult.Yes)
            {
                if (dgv.Rows.Count == 1)
                    MessageBox.Show("Подальше видалення неможливе!");
                else
                {
                    DeleteRow(dgv);
                    RefreshCell();
                }
            }
            else if (dialogResult == DialogResult.No)
            {
                return;
            }
        }

        private void deleteColumnButton_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Ви впевнені, що хочете видалити стовбець? Він може містити важливі дані.",
                "Delete column", MessageBoxButtons.YesNo);
            //удаление столбца
            if (dialogResult == DialogResult.Yes)
            {
                if (dgv.Columns.Count == 1)
                    MessageBox.Show("Подальше видалення неможливе!");
                else
                {
                    DeleteCol(dgv);
                    RefreshCell();
                }
            }
            else if (dialogResult == DialogResult.No) ;
            {
                return;
            }
        }

        private void cellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void toolStripMenuItemAbout_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Ця программа виконує бінарні операціі(+, -, *, /). Також можна застосувати наступні операції: піднесення " +
                "до степеню(^), заперечення(not), порівняння(<, >, ==, <=, >=, <>) та максимум чи мінімум серед кількох елементів" +
                "(mmax(x1, x2, ..., xN), mmin(x1, x2, ..., xN)(N>=1))");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Ви впевнені, цо бажаєте вийти?", "Повідомлення", MessageBoxButtons.YesNo,
                MessageBoxIcon.Exclamation)) 
            {
                //Application.Exit();
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void openStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Зберегти поточний файл?", "Відкрити новий файл", MessageBoxButtons.YesNo);
            if(dialogResult==DialogResult.Yes)
            {
                saveStripMenuItem_Click(sender, e);
            }
            if (openFileDialog1.ShowDialog() != DialogResult.OK)
                return;
            StreamReader sr = new StreamReader(openFileDialog1.FileName);
            dgv.Rows.Clear();
            dgv.Columns.Clear();
            table.Clear();
            int r, c;
            try
            {
                string str = sr.ReadLine();
                r = int.Parse(str);
                str = sr.ReadLine();
                c = int.Parse(str);
                CreateTable(r, c);
                Open(sr, r, c);
                RefreshCell();
            }
            catch
            {
                MessageBox.Show("Помилка при відкритті файлу!");
                dgv.Rows.Clear();
                dgv.Columns.Clear();
                table.Clear();
            }
        }

        private void saveStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                FileStream fs = (FileStream)saveFileDialog1.OpenFile();
                StreamWriter sw = new StreamWriter(fs);
                Save(sw, dgv);
                sw.Close();
                fs.Close();
            }
            RefreshCell();
        }

        public void Save(StreamWriter sw, DataGridView d)
        {
            int row = d.RowCount;
            int col = d.ColumnCount;
            sw.WriteLine(row);
            sw.WriteLine(col);
            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < col; j++)
                {
                    string cellName = SetColNum(j) + (i + 1).ToString();
                    sw.WriteLine(cellName);
                    sw.WriteLine(table[cellName].Exp);
                    int count = table[cellName].CellReference.Count;
                    sw.WriteLine(count);
                    if (count != 0)
                    {
                        foreach(string el in table[cellName].CellReference)
                        {
                            sw.WriteLine(el);
                        }
                    }
                }
            }
        }

        public void Open(StreamReader sr, int row, int col)
        {
            string cellName;
            string expr;
            int count;
            string reference;
            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < col; j++)
                {
                    cellName = sr.ReadLine();
                    expr = sr.ReadLine();
                    table[cellName].Name = cellName;
                    table[cellName].Exp = expr;
                    count = int.Parse(sr.ReadLine());
                    for (int k = 0; k < count; k++)
                    {
                        reference = sr.ReadLine();
                        table[cellName].CellReference.Add(reference);
                    }
                }
            }
        }

        private void CreateTable(int row, int col)
        {
            for (int i = 0; i < col; i++)
            {
                DataGridViewColumn column = new DataGridViewColumn();
                DataGridViewCell cell = new DataGridViewTextBoxCell();
                column.CellTemplate = cell;
                string name = SetColNum(i);
                column.HeaderText = name;
                column.Name = name;
                dgv.Columns.Add(column);
            }

            DataGridViewRow r = new DataGridViewRow();
            for (int i = 0; i < row; i++)
            {
                dgv.Rows.Add();
            }
            SetRowNum(dgv);
            
            Cell _cell;
            for (int j = 0; j < row; j++)
            {
                for (int i = 0; i < col; i++)
                {
                    string cellName = SetColNum(i) + (j + 1).ToString();
                    _cell = new Cell();
                    _cell.Val = "0";
                    _cell.Exp = "";
                    table.Add(cellName, _cell);
                }
            }
        }

        public string SetColNum(int num)
        {
            if (num < 26)
            {
                int n = num + 65;
                return Char.ToString((char)n);
            }
            else
            {
                return Reverse((char)((num % 26) + 65) + SetColNum(num / 26 - 1));
            }
        }

        public void SetRowNum(DataGridView d)
        {
            foreach(DataGridViewRow row in d.Rows)
            {
                row.HeaderCell.Value = String.Format("{0}", row.Index + 1);
            }
        }

        public string Reverse(string str)
        {
            string reversed_str = "";
            for (int i = str.Length - 1; i >= 0; i--) 
            {
                reversed_str += str[i];
            }
            return reversed_str;
        }

        public void AddRow(DataGridView d)
        {
            DataGridViewRow row = new DataGridViewRow();
            d.Rows.Add(row);
            Cell cell;
            for (int i = 0; i < d.ColumnCount; i++)
            {
                string cellName = d.Columns[i].HeaderText + (d.RowCount).ToString();
                cell = new Cell();
                cell.Val = "0";
                cell.Exp = "";
                cell.Name = cellName;
                table.Add(cellName, cell);
            }
            SetRowNum(d);
        }

        public void AddColumn(DataGridView d)
        {
            DataGridViewColumn column = new DataGridViewColumn();
            DataGridViewCell cell = new DataGridViewTextBoxCell();
            column.CellTemplate = cell;
            string name = SetColNum(d.Columns.Count);
            column.Name = name;
            column.HeaderText = name;
            d.Columns.Add(column);
            Cell new_cell;
            for (int i = 0; i < d.RowCount; i++)
            {
                string cellName = name + (i + 1).ToString();
                new_cell = new Cell();
                new_cell.Val = "0";
                new_cell.Exp = "";
                new_cell.Name = cellName;
                table.Add(cellName, new_cell);
            }
        }

        public void DeleteRow(DataGridView d)
        {
            d.Rows.RemoveAt(d.Rows.Count - 1);
            for (int i = 0; i < d.ColumnCount; i++)
            {
                string deletedCell;
                deletedCell = d.Columns[i].HeaderText + (d.Rows.Count + 1).ToString();
                table.Remove(deletedCell);
            }
        }

        public void DeleteCol(DataGridView d)
        {
            string colName = d.Columns[d.Columns.Count - 1].HeaderText;
            for (int i = 0; i < d.RowCount; i++)
            {
                string deletedCell;
                deletedCell = colName + (i + 1).ToString();
                table.Remove(deletedCell);
            }
            d.Columns.RemoveAt(d.Columns.Count - 1);
        }

        private void RefreshCell()
        {
            string errCell = "";
            for (int i = 0; i < dgv.RowCount; i++)
            {
                for (int j = 0; j < dgv.ColumnCount; j++)
                {
                    string cellName = dgv.Columns[j].HeaderText + (i + 1).ToString();
                    if (!table[cellName].WasVisited)
                    {
                        dgv[j, i].Value = "";
                        DFS(cellName, ref errCell);
                    }
                    if (table[cellName].Exp != "")
                    {
                        if (!(table[cellName].Val == "Error" || table[cellName].Val == "ErrorCycle" || table[cellName].Val == "ErrorDivZero"))
                        {
                            dgv[j, i].Value = table[cellName].Val;
                        }
                        else
                        {
                            dgv[j, i].Value = table[cellName].Val;
                        }
                    }
                    else
                    {
                        dgv[j, i].Value = "";
                    }
                }
            }
            foreach (var item in table)
            {
                table[item.Key].WasVisited = false;
            }
            if (errCell != "")
            {
                MessageBox.Show("Помилка в комірці:\n" + errCell);
            }
        }

        private void buttonEnter_Click(object sender, EventArgs e)
        {
            currRow = dgv.CurrentCell.RowIndex;
            currCol = dgv.CurrentCell.ColumnIndex;
            string cellName = dgv.Columns[currCol].HeaderText + (currRow + 1).ToString();
            string str = textBox1.Text;
            table[cellName].Exp = str;
            if (str == "")
            {
                table[cellName].Val = "0";
            }
            table[cellName].CellReference = new List<string>();
            RefreshCell();
            RefreshCell();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                buttonEnter_Click(sender, e);
                this.ActiveControl = dgv;
            }
        }

        private void dgv_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            currRow = dgv.CurrentCell.RowIndex;
            currCol = dgv.CurrentCell.ColumnIndex;
            if (dgv[currCol, currRow].Value != null)
            {
                string cellName = dgv.Columns[currCol].HeaderText + (currRow + 1).ToString();
                textBox1.Text = table[cellName].Exp;
            }
            else
            {
                textBox1.Text = "";
            }
        }

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            currRow = dgv.CurrentCell.RowIndex;
            currCol = dgv.CurrentCell.ColumnIndex;
            string cellName = dgv.Columns[currCol].HeaderText + (currRow + 1).ToString();
            textBox1.Text = table[cellName].Exp;
        }

        public void DFS(string cellName, ref string errCel)
        {
            Calculator calculator = new Calculator(table);
            table[cellName].WasVisited = true;
            foreach(string el in table[cellName].CellReference)
            {
                if (!table.ContainsKey(el))
                {
                    table[cellName].Val = "Error";
                    errCel += "[" + cellName + "]";
                    table[cellName].CellReference.Remove(el);
                    return;
                }
                else if (!table[el].WasVisited) 
                {
                    DFS(el, ref errCel);
                }
            }

            string expr = table[cellName].Exp;
            if (expr != "")
            {
                var res = calculator.Evaluate(expr, cellName, ref errCel);
                if (table[cellName].Val != "Error" && table[cellName].Val != "ErrorCycle" && table[cellName].Val != "ErrorDivZero")
                    table[cellName].Val = res.ToString();
            }
        }
    }
}