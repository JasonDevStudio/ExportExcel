using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportExcel.App_code
{
    /// <summary>
    /// AsposeCell 帮助类
    /// </summary>
    public static class AsposeExcelHelper
    {
        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <typeparam name="T">值类型</typeparam>
        /// <param name="cells">单元格区域</param>
        /// <param name="value">值</param>
        /// <param name="firstRow">区域首行</param>
        /// <param name="firstColumn">区域首列</param>
        /// <param name="totalRows">行数</param>
        /// <param name="totalColumns">列数</param>
        /// <param name="mergeConflict">是否覆盖冲突</param>
        /// <param name="style">样式</param>
        public static void MergeCells<T>(Cells cells, T value, int firstRow, int firstColumn, int totalRows, int totalColumns, bool mergeConflict, Style style = null)
        {
            cells.Merge(firstRow, firstColumn, totalRows, totalColumns, mergeConflict);//合并单元格  
            cells[0, 0].PutValue(value);//填写内容  

            if (style != null)
            {
                cells[0, 0].SetStyle(style);//给单元格关联样式  
            }
        }

        /// <summary>
        /// 创建样式
        /// </summary>
        /// <param name="workbook">Workbook</param>
        /// <param name="backgroundColor">背景色</param>
        /// <param name="foregroundColor">前置色</param>
        /// <param name="fontName">字体名称</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="fontIsBold">是否加粗</param>
        /// <param name="alignment">对齐方式</param>
        /// <param name="lineStyle">边框样式</param>
        public static Style CreateStyle(Workbook workbook, string backgroundColor = "#ffffff", string foregroundColor = "#000000", string fontName = "宋体", int fontSize = 11, bool fontIsBold = false, TextAlignmentType alignment = TextAlignmentType.Center, CellBorderType lineStyle = CellBorderType.None)
        {
            var style = workbook.CreateStyle(); // 新增样式 
            style.HorizontalAlignment = alignment; // 文字居中  
            style.VerticalAlignment = alignment;
            style.Font.Name = fontName; // 文字字体  
            style.Font.Size = fontSize; // 文字大小  
            style.IsLocked = false; // 单元格解锁  
            style.Font.IsBold = fontIsBold; // 粗体  
            style.BackgroundColor = ColorTranslator.FromHtml(backgroundColor);
            style.ForegroundColor = ColorTranslator.FromHtml(foregroundColor); // 设置背景色  
            style.Pattern = BackgroundType.Solid; // 设置背景样式  
            style.IsTextWrapped = true; // 单元格内容自动换行  
            style.Borders[BorderType.LeftBorder].LineStyle = lineStyle; // 应用边界线 左边界线  
            style.Borders[BorderType.RightBorder].LineStyle = lineStyle; // 应用边界线 右边界线  
            style.Borders[BorderType.TopBorder].LineStyle = lineStyle; // 应用边界线 上边界线  
            style.Borders[BorderType.BottomBorder].LineStyle = lineStyle; // 应用边界线 下边界线    
            return style;
        }

        /// <summary>
        /// 添加标题
        /// </summary>
        /// <param name="sheet">Worksheet</param>
        /// <param name="title">标题</param>
        /// <param name="firstRow">首行</param>
        /// <param name="firstColumn">首列</param>
        /// <param name="totalRows">行数</param>
        /// <param name="totalColumns">列数</param>
        /// <param name="style">样式</param>
        private static void SetTitle(this Worksheet sheet, string title, int firstRow, int firstColumn, int totalRows, int totalColumns, Style style)
        {
            sheet.Cells.Merge(firstRow, firstColumn, totalRows, totalColumns);//合并单元格  
            var cell = sheet.Cells[0, 0];
            cell.PutValue(title);
            cell.SetStyle(style);
        }

        /// <summary>
        /// 设置列头
        /// </summary>
        /// <param name="sheet">Worksheet</param>
        /// <param name="name">列头名称</param>
        /// <param name="style">样式</param>
        /// <param name="rowIndex">行号</param>
        /// <param name="columnIndex">列索引</param>
        private static void SetColumnHeader(this Worksheet sheet, string name, Style style, int rowIndex, int columnIndex)
        {
            var cell = sheet.Cells[rowIndex, columnIndex];
            cell.PutValue(name);
            cell.SetStyle(style);
        }

        /// <summary>
        /// 批量设置列头
        /// </summary>
        /// <param name="sheet">Worksheet</param>
        /// <param name="dicProPerties">属性字典</param>
        /// <param name="style">样式</param>
        /// <param name="rowIndex">行号</param>
        /// <param name="columnIndex">列索引</param>
        private static void SetColumnHeaders(this Worksheet sheet, Dictionary<string, string> dicProPerties, Style style, int firstRow, int firstColumn)
        {
            foreach (var item in dicProPerties)
            {
                sheet.SetColumnHeader(item.Value, style, firstRow, firstColumn);
                firstColumn++;
            }
        }

        /// <summary>
        /// 设置值
        /// </summary>
        /// <param name="sheet">Worksheet</param>
        /// <param name="data">DataTable</param>
        /// <param name="dicProPerties">需要写入EXCEL的属性集合</param>
        /// <param name="firstRow">首行</param>
        /// <param name="firstColumn">首列</param>
        private static void SetCellValues(this Worksheet sheet, DataTable data, Dictionary<string, string> dicProPerties, int firstRow, int firstColumn)
        {
            for (int r = 0; r < data.Rows.Count; r++)
            {
                var row = data.Rows[r];

                foreach (var item in dicProPerties)
                {
                    var val = row[item.Key];
                    var cell = sheet.Cells[firstRow, firstColumn];
                    cell.PutValue(val);
                }
            }
        }

        /// <summary>
        /// 设置值
        /// </summary>
        /// <param name="sheet">Worksheet</param>
        /// <param name="data">数据集</param>
        /// <param name="dicProPerties">需要写入EXCEL的属性集合</param>
        /// <param name="firstRow">首行</param>
        /// <param name="firstColumn">首列</param>
        private static void SetCellValues<T>(this Worksheet sheet, List<T> data, Dictionary<string, string> dicProPerties, int firstRow, int firstColumn)
        {
            for (int r = 0; r < data.Count; r++)
            {
                var obj = data[r];
                var t = typeof(T);

                foreach (var item in dicProPerties)
                {
                    var p = t.GetProperty(item.Key);

                    if (p != null)
                    {
                        var val = p.GetValue(obj);

                        if (val != null)
                        {
                            var cell = sheet.Cells[firstRow, firstColumn];
                            cell.PutValue(val);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="data">DataTable 数据源</param>
        /// <param name="dicProPerties">属性结合</param>
        /// <param name="path">导出文件路径</param>
        /// <param name="temp_path">模板路径</param>
        /// <param name="format">保存文件格式</param>
        /// <returns></returns>
        public static bool Export(this DataTable data, Dictionary<string, string> dicProPerties, string path, string temp_path = null, SaveFormat format = SaveFormat.Xlsx)
        {
            var result = false;

            try
            {
                Workbook workbook = null;

                if (temp_path != null && File.Exists(temp_path))
                {
                    workbook = new Workbook(temp_path, new LoadOptions() { MemorySetting = MemorySetting.MemoryPreference }); //工作簿  
                }
                else
                {
                    workbook = new Workbook(FileFormatType.Xlsx); //工作簿  
                }

                var sheet = workbook.Worksheets[0]; //工作表  
                var cells = sheet.Cells;//单元格  

                sheet.SetColumnHeaders(dicProPerties, null, 0, 0);
                sheet.SetCellValues(data, dicProPerties, 1, 0);
                workbook.Save(path, format);

                result = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="data">数据源</param>
        /// <param name="dicProPerties">属性结合</param>
        /// <param name="path">导出文件路径</param>
        /// <param name="temp_path">模板路径</param>
        /// <param name="format">保存文件格式</param>
        /// <returns></returns>
        public static bool Export<T>(this List<T> data, Dictionary<string, string> dicProPerties, string path, string temp_path = null, SaveFormat format = SaveFormat.Xlsx)
        {
            var result = false;

            try
            {
                Workbook workbook = null;

                if (temp_path != null && File.Exists(temp_path))
                {
                    workbook = new Workbook(temp_path, new LoadOptions() { MemorySetting = MemorySetting.MemoryPreference }); //工作簿  
                }
                else
                {
                    workbook = new Workbook(FileFormatType.Xlsx); //工作簿  
                }

                var sheet = workbook.Worksheets[0]; //工作表  
                var cells = sheet.Cells;//单元格  

                sheet.SetColumnHeaders(dicProPerties, null, 0, 0);
                sheet.SetCellValues(data, dicProPerties, 1, 0);
                workbook.Save(path, format);

                result = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }
    }
}
