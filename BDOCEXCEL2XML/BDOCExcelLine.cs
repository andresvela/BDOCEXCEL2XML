using System;
using System.Collections.Generic;
using System.Text;

namespace BDOCEXCEL2XML
{
    public class BDOCExcelLine
    {
        public string dataLevel
        {
            get;
            set;
        }
        public string dataTypeBDOC
        {
            get;
            set;
        }
        public string dataName
        {
            get;
            set;
        }
        public string dataDescription
        {
            get;
            set;
        }
        public string dataIterative
        {
            get;
            set;
        }
        public string dataType
        {
            get;
            set;
        }
        public string dataLong
        {
            get;
            set;
        }
        public string dataXpath
        {
            get;
            set;
        }
        public string dataXMLFlux
        {
            get;
            set;
        }
        public bool reset() {
            this.dataDescription = "";
            this.dataIterative = "";
            this.dataLevel = "";
            this.dataLong = "";
            this.dataName = "";
            this.dataType = "";
            this.dataTypeBDOC = "";
            this.dataXpath = "";
            this.dataXMLFlux = "";
            return true;
        }
    }
}
