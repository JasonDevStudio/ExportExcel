using Aspose.Cells;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
            var work = new Workbook(@"D:\0.xlsx", new LoadOptions() { MemorySetting = MemorySetting.MemoryPreference });
            var sheet = work.Worksheets.FirstOrDefault();
            sheet.Cells.MemorySetting = MemorySetting.MemoryPreference;
            var cells = sheet.Cells;

            var x = 0;
            var y = 0;

            try
            {

                //var table = new DataTable();

                //for (int i = 0; i < 10000; i++)
                //{
                //    table.Columns.Add(String.Format("F_{0}", i));
                //}

                //for (int i = 0; i < 10000; i++)
                //{
                //    var row = table.NewRow();

                //    for (int j = 0; j < 10000; j++)
                //    {
                //        row[j] = string.Format("x={0},y={1}", i, j);
                //    }

                //    table.Rows.Add(row);
                //}

                //MessageBox.Show("Table is success.");


                //sheet.Cells.ImportDataTable(table, true, "A2");

                for (int i = 0; i < 2000; i++)
                {
                    x = i;
                    var row = cells.Rows[i]; 

                    for (int j = 0; j < 15000; j++)
                    {
                        y = j;
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

                        var val = string.Format("x={0},y={1}", i, j);
                        var cell = row[j];
                        cell.PutValue(val);
                    }
                }

                if (File.Exists(@"D:\1.xlsx"))
                {
                    File.Delete(@"D:\1.xlsx");
                }

                work.Save(@"D:\1.xlsx");
                st.Stop();
                var msg = string.Format("Aspose ,数据量10000*10000,耗时[{0}]秒", st.Elapsed.TotalSeconds);
                MessageBox.Show(msg); 
            }
            catch (Exception ex)
            {
                var msg = string.Format("x:{0} y:{1}", x, y);
                MessageBox.Show(msg);
                throw ex;
            }
        }
    }
}
