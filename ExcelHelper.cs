using Npoi.Mapper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelDemo
{
    public class ExcelHelper
    {
        /// <summary>
        /// List转Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="list">数据</param>
        /// <param name="sheetName">表名</param>
        /// <param name="overwrite">true,覆盖单元格，false追加内容(list和创建的excel或excel模板)</param>
        /// <param name="xlsx">true-xlsx，false-xls</param>
        /// <returns>返回文件</returns>
        public static MemoryStream ParseListToExcel<T>(List<T> list, string sheetName = "sheet1", bool overwrite = true, bool xlsx = true) where T : class
        {
            var mapper = new Mapper();
            MemoryStream ms = new MemoryStream();
            mapper.Save<T>(ms, list, sheetName, overwrite, xlsx);
            return ms;
        }
        /// <summary>
        /// Excel转为List
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileStream"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        public static List<T> ParseExcelToList<T>(Stream fileStream, string sheetname = "") where T : class
        {
            List<T> ModelList = new List<T>();
            var mapper = new Mapper(fileStream);
            List<RowInfo<T>> DataList = new List<RowInfo<T>>();
            if (!string.IsNullOrEmpty(sheetname))
            {
                DataList = mapper.Take<T>(sheetname).ToList();
            }
            else
            {
                DataList = mapper.Take<T>().ToList();
            }

            if (DataList != null && DataList.Count > 0)
            {
                foreach (var item in DataList)
                {
                    ModelList.Add(item.Value);
                }
            }
            return ModelList;
        }
    }
}
