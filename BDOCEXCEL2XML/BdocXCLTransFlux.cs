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
    class BdocXCLTransFlux : BdocXCLTransInterface
    {
        private string workSheetName;        
        private Boolean entityIterative;

        private XmlElement childNodestream;
        private XmlElement streamNode2;
        private XmlElement dataStreamNode2;

        public BdocXCLTransFlux(string excelFilePath)
        {
             this.myBDOCExcelLine = new BDOCExcelLine();

            this.xmlDoc = new XmlDocument();
            XmlDeclaration dec = this.xmlDoc.CreateXmlDeclaration("1.0", "ISO-8859-1", null);
            this.xmlDoc.AppendChild(dec);
           
            this.appEXCEL = new Microsoft.Office.Interop.Excel.ApplicationClass();
            // create the workbook object by opening  the excel file.
            workBook = appEXCEL.Workbooks.Open(excelFilePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            // Get The Active Worksheet Using Sheet Name Or Active Sheet
            workSheet = (Worksheet)workBook.ActiveSheet;
            this.workSheetName = workSheet.Name;
            this.excelRange = workSheet.Cells;
            entityIterative = false;

           // throw new System.NotImplementedException();
          
        }

        public override void createStream(BDOCExcelLine excelLine)
        {
            //STREAMS
            if (excelLine.dataLevel == "0")
            {
                XmlElement rootNode = xmlDoc.CreateElement(!(excelLine.dataXpath.Equals("")) ? excelLine.dataXpath : excelLine.dataName);              
                this.xmlDoc.AppendChild(rootNode);
                rootNode.SetAttribute("name", this.workSheetName);
            }

            //first entity who defines the document separator
            if (excelLine.dataLevel == "1")
            {
                XmlNode root = xmlDoc.DocumentElement;

                childNodestream = xmlDoc.CreateElement(!(excelLine.dataXpath.Equals("")) ? excelLine.dataXpath : excelLine.dataName);
                root.AppendChild(childNodestream);                

            }

            // entity struct in excel file
            if (excelLine.dataTypeBDOC == "Entité" && excelLine.dataLevel != "0" && excelLine.dataLevel != "1")
            {
                streamNode = xmlDoc.CreateElement(!(excelLine.dataXpath.Equals("")) ? excelLine.dataXpath : excelLine.dataName);               
                childNodestream.AppendChild(streamNode);

                if (excelLine.dataIterative == "Oui")
                {
                    streamNode2 = xmlDoc.CreateElement(!(excelLine.dataXpath.Equals("")) ? excelLine.dataXpath : excelLine.dataName);
                    childNodestream.AppendChild(streamNode2);
                    entityIterative = true;
                    //childNodestream.AppendChild(streamNode);
                }
                else
                    entityIterative = false;

                //pour les listes
                if (excelLine.dataName.IndexOf("LIST_") == 0)
                {
                    XmlElement dataStreamNode = xmlDoc.CreateElement(!(excelLine.dataXpath.Equals("")) ? excelLine.dataXpath : excelLine.dataName);
                    //dataStreamNode.SetAttribute("name", excelLine.dataName);

                    //XmlElement childNode2 = xmlDoc.CreateElement("xpath");
                    XmlText textNode = xmlDoc.CreateTextNode(excelLine.dataDescription);
                    //If the excel has an specific XPATH value
                    if (excelLine.dataXMLFlux != "")
                        textNode.Value = excelLine.dataXMLFlux;
                    else
                        textNode.Value = excelLine.dataDescription;                          
                    dataStreamNode.AppendChild(textNode);
                    streamNode.AppendChild(dataStreamNode);
                }
            }
            //Data struct in excel file
            if (excelLine.dataTypeBDOC != "Entité" )
            {
                XmlElement dataStreamNode = xmlDoc.CreateElement(!(excelLine.dataXpath.Equals("")) ? excelLine.dataXpath : excelLine.dataName);                
                XmlText textNode = xmlDoc.CreateTextNode(excelLine.dataDescription);
                if (excelLine.dataXMLFlux != "")
                    textNode.Value = excelLine.dataXMLFlux;
                else
                    textNode.Value = excelLine.dataDescription;
                
                dataStreamNode.AppendChild(textNode);
                if (streamNode != null)
                    streamNode.AppendChild(dataStreamNode);
                else
                {
                    throw new Exception(ConfigurationSettings.AppSettings["notEntity1Exception"] + excelLine.dataName + ConfigurationSettings.AppSettings["notEntity2Exception"]);
                }

                if (entityIterative == true) {
                    XmlElement dataStreamNode2 = xmlDoc.CreateElement(!(excelLine.dataXpath.Equals("")) ? excelLine.dataXpath : excelLine.dataName);
                    XmlText textNode2 = xmlDoc.CreateTextNode(excelLine.dataDescription);
                    if (excelLine.dataXMLFlux != "")
                        textNode2.Value = excelLine.dataXMLFlux;
                    else
                        textNode2.Value = excelLine.dataDescription;

                    switch (excelLine.dataType){
                        case "INTEGER":
                            textNode2.Value = textNode2.Value + "2";
                        break;
                        case "STRING":
                            textNode2.Value = textNode2.Value + " 2";
                        break;
                        default:
                            textNode2.Value = textNode2.Value + " 2";
                        break;
                    }                   

                    dataStreamNode2.AppendChild(textNode2);
                    streamNode2.AppendChild(dataStreamNode2);
                }
            }
            //---------------------------------------

        }

        public override void createXML(ToolStripProgressBar MyProgress, string type)
        {
            int index = 1;
            int total = 0;
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
            MyProgress.ToolTipText = total + " registers founded.";
            index = 0;

            try
            {
                while (((Range)workSheet.Cells[rowIndex, colIndex1]).Value2 != null)
                {
                    rowIndex = 2 + index;
                    this.fillBDOCData(myBDOCExcelLine, rowIndex);
                    //XML                      
                    this.createStream(myBDOCExcelLine);
                    MyProgress.PerformStep();
                    index++;
               
                }
                this.FinalXML = this.xmlDoc.InnerXml;
                this.FinalTXT = this.FinalTXTData + "\r\n" + this.FinalTXTEntity;
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
