using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.UserModel.Helpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace JsonToExcel
{
    /// <summary>
    /// 公共工具类
    /// </summary>
    public class CommonHelper
    {

        //获取列名委托方法
        public delegate string GetColumnName(string columnName);

        #region 导入导出Excel相关

        /// <summary>
        /// 将泛类型集合List类转换成DataTable
        /// </summary>
        /// <param name="list">泛类型集合</param>
        /// <returns>返回转换后的DataTable</returns>
        public static DataTable ListToDataTable<T>(List<T> entitys)
        {
            //生成DataTable的structure
            var dt = new DataTable();
            try
            {
                //检查泛型实体是否为空
                if (entitys == null || entitys.Count < 1)
                {
                    return dt;
                }
                //取出第一个实体的所有Propertie
                var entityType = entitys[0].GetType();
                var entityProperties = entityType.GetProperties();
                for (var i = 0; i < entityProperties.Length; i++)
                {
                    dt.Columns.Add(entityProperties[i].Name);
                }
                //将所有entity添加到DataTable中
                foreach (object entity in entitys)
                {
                    //检查所有的的实体都为同一类型
                    if (entity.GetType() != entityType)
                    {
                        throw new Exception("要转换的集合元素类型不一致");
                    }
                    var entityValues = new object[entityProperties.Length];
                    for (var i = 0; i < entityProperties.Length; i++)
                    {
                        entityValues[i] = entityProperties[i].GetValue(entity, null);
                    }
                    dt.Rows.Add(entityValues);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return dt;
        }

        /// <summary>
        /// 将dataTable转换为Excel字节流
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="getColumnName"></param>
        /// <returns></returns>
        /// <summary>
        /// NPOI简单Demo，快速入门代码
        /// </summary>
        /// <param name="dtSource"></param>
        /// <param name="strFileName"></param>
        /// <remarks>NPOI认为Excel的第一个单元格是：(0，0)</remarks>
        public static void ExportEasy(DataTable dtSource, string strFileName)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet();

            //填充表头
            IRow dataRow = sheet.CreateRow(0);
            foreach (DataColumn column in dtSource.Columns)
            {
                dataRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
            }


            //填充内容
            for (int i = 0; i < dtSource.Rows.Count; i++)
            {
                dataRow = sheet.CreateRow(i + 1);
                for (int j = 0; j < dtSource.Columns.Count; j++)
                {
                    dataRow.CreateCell(j).SetCellValue(dtSource.Rows[i][j].ToString());
                }
            }


            //保存
            using (MemoryStream ms = new MemoryStream())
            {
                using (FileStream fs = new FileStream(strFileName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }
            }
        }
    }

    #endregion
}

