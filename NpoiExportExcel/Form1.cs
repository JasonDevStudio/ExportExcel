using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NpoiExportExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var st = new Stopwatch();
            st.Restart();

            var book = new XSSFWorkbook();
            var sheet = book.CreateSheet("001");

            for (int i = 0; i < 10000; i++)
            {
                var row = sheet.CreateRow(i);

                for (int j = 0; j < 1000; j++)
                {
                    var cell = row.CreateCell(j);
                    cell.SetCellValue(Guid.NewGuid().ToString());
                }
            }

            if (File.Exists(@"D:\3.xlsx"))
            {
                File.Delete(@"D:\3.xlsx");
            }

            using (var fs = new FileStream(@"D:\3.xlsx",FileMode.Create,FileAccess.Write))
            {
                book.Write(fs);
                fs.Close();
            }

            st.Stop();
            var msg = string.Format("Npoi,数据量10000*1000,耗时[{0}]秒", st.Elapsed.TotalSeconds);
            MessageBox.Show(msg);
            Clipboard.SetDataObject(msg, true);
        }
    }
}
