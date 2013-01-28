using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Configuration;

namespace BDOCEXCEL2XML
{
    class BDOCXCLTransTxt : BdocXCLTransInterface
    {

        public BDOCXCLTransTxt(string excelFilePath)
        {

            this.myBDOCExcelLine = new BDOCExcelLine();
            this.appEXCEL = new Microsoft.Office.Interop.Excel.ApplicationClass();
            // create the workbook object by opening  the excel file.
            workBook = appEXCEL.Workbooks.Open(excelFilePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            // Get The Active Worksheet Using Sheet Name Or Active Sheet
            workSheet = (Worksheet)workBook.ActiveSheet;
            this.excelRange = workSheet.Cells;
            //************ TXT init
            FinalTXTData = "";
            FinalTXTEntity = "";
            FinalTXTStream = "";

           // throw new System.NotImplementedException();
          
        }

        public override void createData(BDOCExcelLine excelLine)
        {
            //Pour les donnees
            if (excelLine.dataTypeBDOC != "Entité")
            {
                FinalTXTData += "*" + "\r\n";
                FinalTXTData += excelLine.dataName + "\r\n";
                FinalTXTData += excelLine.dataDescription + "\r\n";
                FinalTXTData += excelLine.dataLong + "\r\n";
                switch (excelLine.dataType)
                {
                    case "STRING":
                        FinalTXTData += "0";
                        break;
                    case "CHAR":
                        FinalTXTData += "1";
                        break;
                    case "INTEGER":
                        FinalTXTData += "2";
                        break;
                    case "REAL":
                        FinalTXTData += "3";
                        break;
                }
                FinalTXTData += "\r\n";
            }
        }

        public override void createStream(BDOCExcelLine excelLine)
        {
               //pour les streamses
            //premiere entite qui definira le separateur de document
            if (excelLine.dataLevel == "1")
            {
                FinalTXTStream += "*" + "\r\n"; // 1 flux separator
                FinalTXTStream += excelLine.dataName + "\r\n"; // 2 flux name
                FinalTXTStream += excelLine.dataDescription + "\r\n"; // 3 flux descr
                FinalTXTStream += "0" + "\r\n"; // 4 Stream type , only 0 available
                FinalTXTStream += "0" + "\r\n"; // 5 Line type , only 0 available
                FinalTXTStream += "+0 +3" + "\r\n"; // 6 TypeID ,see admin doc, page 69
                FinalTXTStream += "\r\n"; // 7 stream file path, 
                FinalTXTStream += "0" + "\r\n"; // 8 Reserved , only 0 available
            }

            if (excelLine.dataTypeBDOC == "Entité" && excelLine.dataLevel != "0" && excelLine.dataLevel != "1")
            {
                FinalTXTStream += "*E" + "\r\n"; // 9 entity separator
                FinalTXTStream += excelLine.dataName + "\r\n";// 10 entity name
                FinalTXTStream += excelLine.dataDescription + "\r\n";// 11 entity separator                
                FinalTXTStream += "0" + "\r\n";// 12 entity type fonctionnel
                FinalTXTStream += (excelLine.dataIterative == "Non") ? "0" + "\r\n" : "1" + "\r\n"; // 13 entity type iterative                
                FinalTXTStream += "0" + "\r\n";// 14 reserved
                FinalTXTStream += "001" + "\r\n"; // 15 Id, CORRECT IT!

                FinalTXTStream += "2" + "\r\n"; // 16 Debut document
                FinalTXTStream += "0" + "\r\n"; // 17 Reserved DOC1
                FinalTXTStream += "3" + "\r\n"; // 18 Reserved DOC1

                /* //pour les listes
                if (excelLine.dataName.IndexOf("LIST_") == 0)
                {
                    XmlElement dataStreamNode = xmlDoc.CreateElement("dataNode");
                    dataStreamNode.SetAttribute("name", excelLine.dataName);

                    XmlElement childNode2 = xmlDoc.CreateElement("xpath");
                    textNode = xmlDoc.CreateTextNode(excelLine.dataDescription);
                    if (excelLine.dataName.Length > 18)
                        textNode.Value = "descendant-or-self::OCCURRENCES_" + excelLine.dataName.Substring(0, 18);
                    else
                        textNode.Value = "descendant-or-self::OCCURRENCES_" + excelLine.dataName;
                    childNode2.AppendChild(textNode);
                    dataStreamNode.AppendChild(childNode2);
                    streamNode.AppendChild(dataStreamNode);
                }
                * */
            }
            if (excelLine.dataTypeBDOC != "Entité")
            {
                FinalTXTStream += "*D" + "\r\n"; // data separator
                FinalTXTStream += excelLine.dataName + "\r\n"; // Data name
                FinalTXTStream += excelLine.dataDescription + "\r\n"; // 21 data description
                FinalTXTStream += excelLine.dataLong + "\r\n"; // 22 long data
                switch (excelLine.dataType)
                {
                    case "STRING":
                        FinalTXTStream += "0" + "\r\n";
                        break;
                    case "CHAR":
                        FinalTXTStream += "1" + "\r\n";
                        break;
                    case "INTEGER":
                        FinalTXTStream += "2" + "\r\n";
                        break;
                    case "REAL":
                        FinalTXTStream += "3" + "\r\n";
                        break;
                } // 23 type donnée
                FinalTXTStream += "+1 >#<" + "\r\n"; // 24 delimiteur
                FinalTXTStream += "0" + "\r\n"; // 25 DOC1 Filier
                FinalTXTStream += "0" + "\r\n"; //26 reservée                                
            }
            //---------------------------------------
        }          
            
        public override void createEntity(BDOCExcelLine excelLine)
        {
            //pour les entities           
            if (excelLine.dataTypeBDOC == "Entité" && excelLine.dataName != "FLUX" && excelLine.dataLevel != "0")
            {
                if (this.lastObjectInserted == "ENTITY")
                    FinalTXTEntity = FinalTXTEntity.Substring(0, FinalTXTEntity.LastIndexOf("*"));

                FinalTXTEntity += "*" + "\r\n";
                FinalTXTEntity += excelLine.dataName + "\r\n";
                FinalTXTEntity += excelLine.dataDescription + "\r\n";
                FinalTXTEntity += "0" + "\r\n"; //functional entity
                FinalTXTEntity += ((excelLine.dataIterative == "Non") ? "0" : "1") + "\r\n";
                FinalTXTEntity += "0" + "\r\n"; //gestion d'habilitation 

                this.lastObjectInserted = "ENTITY";
                if (excelLine.dataName.IndexOf("LIST_") == 0)
                {
                    if (excelLine.dataName.Length > 18)
                        FinalTXTEntity += "OCCURRENCES_" + excelLine.dataName.Substring(0, 18) + "\r\n";
                    else
                        FinalTXTEntity += "OCCURRENCES_" + excelLine.dataName + "\r\n";
                }
            }
            if (excelLine.dataTypeBDOC != "Entité")
            {
                FinalTXTEntity += excelLine.dataName + "\r\n";
                this.lastObjectInserted = "DATA";
            }
            //---------------------------------------
        }

        /// <summary>
        /// Fonction qui construira un XML à partir d'un Excel donné.
        /// </summary>
        public override void createXML(ToolStripProgressBar MyProgress)
        {
            int index = 1;
            this.total = 0;
            object rowIndex = 2;
            object colIndex1 = 1;
            object colIndex2 = 2;
            object colIndex3 = 3;
            object colIndex4 = 4;
            object colIndex5 = 5;
            object colIndex6 = 6;
            object colIndex7 = 7;

            //get total lines
            MyProgress.Value = 50;
            while (((Range)workSheet.Cells[index, colIndex1]).Value2 != null)
            {
                total = index++;
            }
            MyProgress.Maximum = total;
            MyProgress.Step = 1;
            MyProgress.Value = 0;
            
            index = 0;

            try
            {
                while (((Range)workSheet.Cells[rowIndex, 1]).Value2 != null)
                {
                    rowIndex = 2 + index;                   
                    this.fillBDOCData(myBDOCExcelLine, rowIndex);

                    //XML
                    this.createData(myBDOCExcelLine);
                    this.createEntity(myBDOCExcelLine);
                    this.createStream(myBDOCExcelLine);

                    MyProgress.PerformStep();

                    index++;                    
                }               

                //this.FinalTXT = this.FinalTXTData + "\r\n" + this.FinalTXTEntity;

                this.appEXCEL.Quit();

            }
            catch (Exception ex)
            {
                appEXCEL.Quit();
                Console.WriteLine(ex.Message);
                throw ex;
            }


        }


    }
}
