using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using System.Diagnostics;

namespace EXCEL_DEV
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string path = string.Format(@"{0}\DATA\单元格.xlsx",Application.StartupPath);
            excel.LoadDocument(path);
        }

        private void btn_Click(object sender, EventArgs e)
        {
            IWorkbook book = excel.Document;
            Worksheet sht = book.Worksheets[0];
            for (int i = 0; i < 10;i++ )
            {
                Cell cell = sht.Cells[i, 0];
                Debug.Print("{0}:{1}",cell.DisplayText,cell.NumberFormat);

            }
        }

        private void btnSet_Click(object sender, EventArgs e)
        {
            IWorkbook book = excel.Document;
            Worksheet sht = book.Worksheets[0];
            for (int i = 0; i < 10; i++)
            {
                Cell cell = sht.Cells[i, 0];
                cell.NumberFormat = "0.000";
                Debug.Print("{0}:{1}", cell.DisplayText, cell.NumberFormat);

            }
        }
    }
}
