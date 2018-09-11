using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Loigc;
using Microsoft.AspNetCore.Mvc;

namespace NPOIFastExport.Controllers
{
    public class DemoController : Controller
    {

        public IActionResult Index()
        {

            return View();
        }
        /// <summary>
        /// 制造DataTable
        /// </summary>
        /// <returns></returns>
        private DataTable GetData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("姓名");
            dt.Columns.Add("性别");
            dt.Columns.Add("爱好");
            dt.Columns.Add("住址");
            dt.Columns.Add("星座");
            for (int i = 0; i < 10; i++)
            {
                DataRow row = dt.NewRow();
                row["姓名"] = "阿塔尼斯";
                row["性别"] = "男";
                row["爱好"] = "装逼";
                row["住址"] = "艾尔";
                row["星座"] = "不知道";
                dt.Rows.Add(row);
            }
            return dt;
        }

        /// <summary>
        /// 导出Excel(不带表头)
        /// </summary>
        /// <returns></returns>
        public IActionResult ExportExcel()
        {
            var data = GetData();
            var fileName = $"测试导出";
            var excelwork = new ExcelWorker();
            var bytes = excelwork.ExportDataTable(data, fileName);
            return File(bytes, "application/x-xls", fileName + ".xls");
        }


        /// <summary>
        /// 导出Excel(带表头)
        /// </summary>
        /// <returns></returns>
        public IActionResult ExportExcelByTitle(string name=null)
        {
            var data = GetData();
            var fileName = $"测试导出";
            var excelwork = new ExcelWorker();
            var bytes = excelwork.ExportDataTable(data, fileName,name==null?"测试表头":name);
            return File(bytes, "application/x-xls", fileName + ".xls");
        }
    }
}