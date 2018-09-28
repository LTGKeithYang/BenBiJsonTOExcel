using System;
using System.Collections.Generic;
using System.Text;

namespace JsonToExcel
{
    class Model
    {
        public int id { get; set; }
        public string periodsName { get; set; }
        public string trainPlace { get; set; }
        public long startTime { get; set; }
        public long endTime { get; set; }
        public long trainStartTime { get; set; }
        public long trainEndTime { get; set; }
        public int limitNum { get; set; }
        public long createDate { get; set; }
        public long updateDate { get; set; }
        public int userNum { get; set; }
        public int passmark { get; set; }
        public int excellentPoints { get; set; }
        public int trainType { get; set; }
    }
}
