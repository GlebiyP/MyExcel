using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyExcel
{
    public class Cell : DataGridViewTextBoxCell
    {
        string _value;
        string name;
        string exp;
        bool wasVisited = false;
        List<string> cellReference;
        public Cell()
        {
            name = "A1";
            exp = "";
            _value = "0";
            cellReference = new List<string>();
            wasVisited = false;
        }

        public string Val
        {
            get { return _value; }
            set { _value = value; }
        }

        public string Exp
        {
            get { return exp; }
            set { exp = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public List<string> CellReference
        {
            get { return cellReference; }
            set { cellReference = value; }
        }

        public bool WasVisited
        {
            get { return wasVisited; }
            set { wasVisited = value; }
        }
    }
}
