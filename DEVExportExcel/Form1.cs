using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet.Model;
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

namespace DEVExportExcel
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
            var work = new DocumentModel();
            using (var fs = new FileStream(@"D:\0.xlsx", FileMode.Open, FileAccess.Read))
            {
                work.LoadDocument(fs, DocumentFormat.Xlsx, null);
            }

            var sheet = work.Sheets.FirstOrDefault();

            for (int i = 0; i < 1000; i++)
            {
                for (int j = 0; j < 1000; j++)
                {
                    //var str = new StringBuilder();
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());

                    var post = new CellPosition(i, j);
                    //var comment = sheet.CreateComment(post, " ");
                    //comment.Shape.FillColor = Color.Yellow;
                    //comment.SetPlainText(str.ToString());
                    //comment.Shape.ClientData.Anchor.SetFrom(0, j, 0, i, 0);
                    //comment.Shape.ClientData.Anchor.SetTo(0, j + 10, 0, i + 15, 0);
                    sheet[post].SetValue(new VariantValue() { InlineTextValue = Guid.NewGuid().ToString() });
                }
            }

            if (File.Exists(@"D:\2.xlsx"))
            {
                File.Delete(@"D:\2.xlsx");
            }

            using (var fs = new FileStream(@"D:\2.xlsx", FileMode.Create, FileAccess.Write))
            {
                work.SaveDocument(fs, DocumentFormat.Xlsx, null);
                fs.Close(); 
            }

            st.Stop();
            var msg = string.Format("DevExpress,耗时[{0}]秒", st.Elapsed.TotalSeconds);
            MessageBox.Show(msg);
            Clipboard.SetDataObject(msg, true);
        }
    }
}
