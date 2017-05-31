using Aspose.Cells;
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

namespace ExportExcel
{
    public partial class FrmAspose : Form
    {
        public FrmAspose()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var st = new Stopwatch();
            st.Restart();
            var work = new Workbook(FileFormatType.Xlsx);
            var sheet = work.Worksheets.FirstOrDefault();

            for (int i = 0; i < 10000; i++)
            {
                for (int j = 0; j < 10000; j++)
                {
                    //var str = new StringBuilder();
                    //str.AppendLine("<html");
                    //str.AppendLine("<body style='background - color:yellow'>");
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine(Guid.NewGuid().ToString());
                    //str.AppendLine("</ body >");
                    //str.AppendLine("</ html >");

                    //sheet.Comments.Add(i, j);
                    //var comment = sheet.Comments[i, j];
                    //comment.AutoSize = true;
                    //comment.CommentShape.HtmlText = str.ToString();
                    //comment.HtmlNote = str.ToString();

                    var cell = sheet.Cells[i, j];
                    cell.Value = Guid.NewGuid().ToString();

                }
            }

            if (File.Exists(@"D:\1.xlsx"))
            {
                File.Delete(@"D:\1.xlsx");
            }

            work.Save(@"D:\1.xlsx");
            st.Stop();
            var msg = string.Format("Aspose,数据量10000*10000,耗时[{0}]秒", st.Elapsed.TotalSeconds);
            MessageBox.Show(msg);
            Clipboard.SetDataObject(msg, true);
        }
    }
}
