using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace Loigc
{
    public class ExcelWorker
    {
        readonly int EXCEL03_MaxRow = 65535;

        /// <summary>
        /// 把一个DataTable导出成一个Excel.只有一个Sheet
        /// </summary>
        /// <param name="dt">要导出的DataTable</param>
        /// <param name="sheetName">sheet的名字</param>
        /// <param name="tiltle">把表格顶部的第一行设置为表头，并指定表头的内容，如果不需要表头的话传null</param>
        /// <returns></returns>
        public byte[] ExportDataTable(DataTable dt, string sheetName, string tiltle = null)
        {
            IWorkbook book = new HSSFWorkbook();
            if (dt.Rows.Count < EXCEL03_MaxRow)
                DataWrite2Sheet(dt, 0, dt.Rows.Count - 1, book, sheetName, tiltle);
            else
            {
                int page = dt.Rows.Count / EXCEL03_MaxRow;
                for (int i = 0; i < page; i++)
                {
                    int start = i * EXCEL03_MaxRow;
                    int end = (i * EXCEL03_MaxRow) + EXCEL03_MaxRow - 1;
                    DataWrite2Sheet(dt, start, end, book, sheetName + i.ToString(), tiltle);
                }
                int lastPageItemCount = dt.Rows.Count % EXCEL03_MaxRow;
                DataWrite2Sheet(dt, dt.Rows.Count - lastPageItemCount, lastPageItemCount, book, sheetName + page.ToString());
            }
            using (MemoryStream ms = new MemoryStream())
            {
                book.Write(ms);
                return ms.ToArray();
            }
        }

        /// <summary>
        /// 在一个Excel中导出多个Sheet,适用于每个Sheet都有自己的表头
        /// </summary>
        /// <param name="dataTables">要传入的DataTable</param>
        /// <param name="sheetName">这个Excel内Sheet名字集合，和DataTable的顺序对应</param>
        /// <param name="tiltle">每个sheet表格的表头和sheet顺序对应</param>
        /// <returns></returns>
        public byte[] ExportDataTableAll(List<DataTable> dataTables, List<string> sheetName, List<string> tiltle = null)
        {
            IWorkbook book = new HSSFWorkbook();
            int j = 0;
            if (tiltle != null && tiltle.Count > 0 && tiltle.Count == dataTables.Count)
            {
                foreach (var dt in dataTables)
                {

                    if (dt.Rows.Count < EXCEL03_MaxRow)
                        DataWrite2Sheet(dt, 0, dt.Rows.Count - 1, book, sheetName[j], tiltle[j]);
                    else
                    {
                        int page = dt.Rows.Count / EXCEL03_MaxRow;
                        for (int i = 0; i < page; i++)
                        {
                            int start = i * EXCEL03_MaxRow;
                            int end = (i * EXCEL03_MaxRow) + EXCEL03_MaxRow - 1;
                            DataWrite2Sheet(dt, start, end, book, sheetName[j] + i.ToString(), tiltle[j]);
                        }
                        int lastPageItemCount = dt.Rows.Count % EXCEL03_MaxRow;
                        DataWrite2Sheet(dt, dt.Rows.Count - lastPageItemCount, lastPageItemCount, book, sheetName[j] + page.ToString(), sheetName[j]);
                    }
                    j++;
                }
            }
            using (MemoryStream ms = new MemoryStream())
            {
                book.Write(ms);
                return ms.ToArray();
            }
        }

        /// <summary>
        /// 在一个Excel中导出多个Sheet,适用于每个Sheet都没有表头
        /// </summary>
        /// <param name="dataTables">要传入的DataTable</param>
        /// <param name="sheetName">这个Excel内Sheet表头名字集合，和DataTable的顺序对应</param>
        /// <returns></returns>
        public byte[] ExportDataTableAll(List<DataTable> dataTables, List<string> sheetName)
        {
            IWorkbook book = new HSSFWorkbook();
            int j = 0;
            foreach (var dt in dataTables)
            {

                if (dt.Rows.Count < EXCEL03_MaxRow)
                    DataWrite2Sheet(dt, 0, dt.Rows.Count - 1, book, sheetName[j]);
                else
                {
                    int page = dt.Rows.Count / EXCEL03_MaxRow;
                    for (int i = 0; i < page; i++)
                    {
                        int start = i * EXCEL03_MaxRow;
                        int end = (i * EXCEL03_MaxRow) + EXCEL03_MaxRow - 1;
                        DataWrite2Sheet(dt, start, end, book, sheetName[j] + i.ToString());
                    }
                    int lastPageItemCount = dt.Rows.Count % EXCEL03_MaxRow;
                    DataWrite2Sheet(dt, dt.Rows.Count - lastPageItemCount, lastPageItemCount, book, sheetName[j] + page.ToString(), sheetName[j]);
                }
                j++;

            }
            using (MemoryStream ms = new MemoryStream())
            {
                book.Write(ms);
                return ms.ToArray();
            }
        }

        /// <summary>
        /// 生成Excel
        /// </summary>
        /// <param name="dt">转化的Datatable</param>
        /// <param name="startRow">从第几行开始写</param>
        /// <param name="endRow"></param>
        /// <param name="book">NOPI wokerbook对象</param>
        /// <param name="sheetName">在EXCEL底部显示的sheet名字</param>
        /// <param name="tiltle">表格顶部第一行设置为表头</param>
        private void DataWrite2Sheet(DataTable dt, int startRow, int endRow, IWorkbook book, string sheetName, string tiltle = null)
        {
            if (!string.IsNullOrEmpty(tiltle))
            {
                ISheet sheet = book.CreateSheet(sheetName);
                IRow header = sheet.CreateRow(0);
                ICell cell1 = header.CreateCell(0);
                if (!string.IsNullOrEmpty(tiltle))
                {
                    cell1.SetCellValue(tiltle);
                    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, dt.Columns.Count));
                }
                IRow header2 = sheet.CreateRow(1);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    ICell cell = header2.CreateCell(i);
                    string val = dt.Columns[i].Caption ?? dt.Columns[i].ColumnName;
                    cell.SetCellValue(val);
                }
                int rowIndex = 2;
                for (int i = 0; i <= endRow; i++)
                {
                    DataRow dtRow = dt.Rows[i];
                    IRow excelRow = sheet.CreateRow(rowIndex++);
                    for (int j = 0; j < dtRow.ItemArray.Length; j++)
                    {
                        excelRow.CreateCell(j).SetCellValue(dtRow[j].ToString());
                    }
                }
            }
            else
            {
                ISheet sheet = book.CreateSheet(sheetName);
                IRow header = sheet.CreateRow(0);
                ICell cell1 = header.CreateCell(0);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    ICell cell = header.CreateCell(i);
                    string val = dt.Columns[i].Caption ?? dt.Columns[i].ColumnName;
                    cell.SetCellValue(val);
                }
                int rowIndex = 1;
                for (int i = startRow; i <= endRow; i++)
                {
                    DataRow dtRow = dt.Rows[i];
                    IRow excelRow = sheet.CreateRow(rowIndex++);
                    for (int j = 0; j < dtRow.ItemArray.Length; j++)
                    {
                        excelRow.CreateCell(j).SetCellValue(dtRow[j].ToString());
                    }
                }
            }
        }

        /// <summary>
        /// 导入Excel,把Excel转成DataTable  emmmmm这个封装好像有毛病，只支持xlsx
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="isColumnName"></param>
        /// <returns>第一行是否是表头</returns>
        public static DataTable ExcelToDataTable(string filePath, bool isColumnName)
        {
            DataTable dataTable = null;
            FileStream fs = null;
            DataColumn column = null;
            DataRow dataRow = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            IRow row = null;
            ICell cell = null;
            int startRow = 0;
            try
            {
                using (fs = File.OpenRead(filePath))
                {
                    if (filePath.IndexOf(".xlsx") > 0)
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    else if (filePath.IndexOf(".xls") > 0)
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    if (workbook != null)
                    {
                        sheet = workbook.GetSheetAt(0);
                        dataTable = new DataTable();
                        if (sheet != null)
                        {
                            int rowcCount = sheet.LastRowNum;
                            if (rowcCount > 0)
                            {
                                IRow firstRow = sheet.GetRow(0);
                                int cellCount = firstRow.LastCellNum;

                                if (isColumnName)
                                {
                                    startRow = 1;
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        cell = firstRow.GetCell(i);
                                        if (cell != null)
                                        {
                                            if (cell.StringCellValue != null)
                                            {
                                                column = new DataColumn(cell.StringCellValue);
                                                dataTable.Columns.Add(column);
                                            }
                                        }

                                    }
                                }
                                else
                                {
                                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                                    {
                                        column = new DataColumn("colum" + (i + 1));
                                        dataTable.Columns.Add(column);
                                    }
                                }
                                for (int i = startRow; i <= rowcCount; ++i)
                                {
                                    row = sheet.GetRow(i);
                                    if (row == null)
                                    {
                                        continue;
                                    }
                                    dataRow = dataTable.NewRow();
                                    for (int j = row.FirstCellNum; j < cellCount; ++j)
                                    {
                                        cell = row.GetCell(j);
                                        if (cell == null)
                                        {
                                            dataRow[j] = "";
                                        }
                                        else
                                        {
                                            switch (cell.CellType)
                                            {
                                                case CellType.Blank:
                                                    dataRow[j] = "";
                                                    break;
                                                case CellType.Numeric:
                                                    short format = cell.CellStyle.DataFormat;
                                                    if (format == 14 || format == 31 || format == 57 || format == 58)
                                                        dataRow[j] = cell.DateCellValue;
                                                    else
                                                        dataRow[j] = cell.NumericCellValue;
                                                    break;
                                                case CellType.String:
                                                    dataRow[j] = cell.StringCellValue;
                                                    break;
                                            }
                                        }
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                        }
                    }
                }
                return dataTable;
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                return null;
            }
        }
    }
}
