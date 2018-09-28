using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;


namespace JsonToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string jsonfile = @"C:\Users\KeithYang\source\repos\JsonToExcel\JsonToExcel\1.txt";
            using (System.IO.StreamReader file = System.IO.File.OpenText(jsonfile))
            {
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    JObject o = (JObject)JToken.ReadFrom(reader);
                    DataTable table = new DataTable();
                    table.Columns.Add("id");
                    table.Columns.Add("periodsName");
                    table.Columns.Add("trainPlace");
                    table.Columns.Add("startTime");
                    table.Columns.Add("endTime");
                    table.Columns.Add("trainStartTime");
                    table.Columns.Add("trainEndTime");
                    table.Columns.Add("limitNum");
                    table.Columns.Add("userNum");
                    table.Columns.Add("createDate");
                    table.Columns.Add("updateDate");
                    table.Columns.Add("passmark");
                    table.Columns.Add("excellentPoints");
                    table.Columns.Add("trainType");

                    Console.WriteLine(o.ToString());
                    JArray a = (JArray)o["list"];
                    List<Model> lm=new List<Model>();
                    foreach (var i in a)
                    {
                        lm.Add(i.ToObject<Model>());
                    }
                    foreach (var model in lm)
                    {
                        DataRow dr = table.NewRow();
                        dr["id"] = model.id;
                        dr["periodsName"] = model.periodsName;
                        dr["trainPlace"] = model.trainPlace;
                        dr["startTime"] = ToCurrentTimeZone(model.startTime);
                        dr["endTime"] = ToCurrentTimeZone(model.endTime);
                        dr["trainStartTime"] = ToCurrentTimeZone(model.trainStartTime);
                        dr["trainEndTime"] = ToCurrentTimeZone(model.trainEndTime);
                        dr["limitNum"] = model.limitNum;
                        dr["userNum"] = model.userNum;
                        dr["createDate"] = ToCurrentTimeZone(model.createDate);
                        dr["updateDate"] = ToCurrentTimeZone(model.updateDate);
                        dr["passmark"] = model.passmark;
                        dr["excellentPoints"] = model.excellentPoints;
                        dr["trainType"] = model.trainType;
                        table.Rows.Add(dr);
                    }
                    //var b=CommonHelper.ListToDataTable(lm);
                    CommonHelper.ExportEasy(table,"infomation.xls");

                }

            }
        }

        public static DateTime ToCurrentTimeZone(long timeStamp)
        {
            System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1)); // 当地时区
            DateTime dt = startTime.AddMilliseconds(timeStamp);
            return dt;
        }

        public static string KeepColumnName(string colName)
        {
            return colName;
        }


    }
}

