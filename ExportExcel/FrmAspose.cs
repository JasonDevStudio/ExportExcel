using Aspose.Cells;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExportExcel.App_code;
using System.Collections.Generic;

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

        private void button2_Click(object sender, EventArgs e)
        {
            var table = new DataTable();
            var dics = new Dictionary<string, string>();
            table.Columns.Add(new DataColumn() { ColumnName = "ID", Caption = "编号" });
            table.Columns.Add(new DataColumn() { ColumnName = "NAME", Caption = "名称" });
            table.Columns.Add(new DataColumn() { ColumnName = "VALUE", Caption = "值" });
            table.Columns.Add(new DataColumn() { ColumnName = "TEMP", Caption = "温度" });
            table.Columns.Add(new DataColumn() { ColumnName = "CORNER", Caption = "工艺" });
            table.Columns.Add(new DataColumn() { ColumnName = "TESTER", Caption = "机台" });
            table.Columns.Add(new DataColumn() { ColumnName = "DEVICENO", Caption = "芯片编号" });
            table.Columns.Add(new DataColumn() { ColumnName = "SUBNAME", Caption = "Sub Name" });

            foreach (DataColumn item in table.Columns)
            {
                dics.Add(item.ColumnName, item.Caption);
            }

            for (int i = 0; i < 1000; i++)
            {
                var row = table.NewRow();
                row["ID"] = i;
                row["NAME"] = string.Format("名: {0} ", i);
                row["VALUE"] = i + 1;
                row["DEVICENO"] = i + 1;
                row["TEMP"] = string.Format("{0}C", i);
                row["CORNER"] = "TT";
                row["TESTER"] = "DDFGSHGKJDFGKDJHSDGF";
                row["SUBNAME"] = Guid.NewGuid().ToString();
                table.Rows.Add(row);
            }

            table.Export(dics, @"D:\Temp\1.xlsx", format: SaveFormat.Xlsx);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var entyties = new List<TestEntity>();
            var dics = new Dictionary<string, string>();
            dics.Add("ID", "编号");
            dics.Add("NAME", "名称");
            dics.Add("VALUE", "值");
            dics.Add("TEMP", "温度");
            dics.Add("CORNER", "工艺");
            dics.Add("TESTER", "机台");

            for (int i = 0; i < 10; i++)
            {
                entyties.Add(new TestEntity()
                {
                    ID = i,
                    NAME = string.Format("名: {0} ", i),
                    VALUE = "" + i + 1,
                    TEMP = string.Format("{0}C", i),
                    CORNER = "TTEE",
                    TESTER = "GGDERHSDFDSFFSF"
                });
            }

            entyties.Export(dics, @"D:\Temp\2.xlsx", format: SaveFormat.Xlsx);
        }
    }

    public class TestEntity
    {
        public int ID { get; set; }
        public string NAME { get; set; }
        public string VALUE { get; set; }
        public string TEMP { get; set; }
        public string CORNER { get; set; }
        public string TESTER { get; set; }
    }
}
