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
    class BdocXCLTransImport : BdocXCLTransInterface
    {
        public BdocXCLTransImport(string excelFilePath)
        {
             this.myBDOCExcelLine = new BDOCExcelLine();

            this.xmlDoc = new XmlDocument();
            XmlDeclaration dec = this.xmlDoc.CreateXmlDeclaration("1.0", "ISO-8859-1", null);
            this.xmlDoc.AppendChild(dec);
            XmlElement rootNode = xmlDoc.CreateElement("import");
            rootNode.SetAttribute("xmlns", "http://www.bdoc.com");
            rootNode.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");                        
            this.xmlDoc.AppendChild(rootNode);

            XmlNode root = xmlDoc.DocumentElement;
            
            XmlElement childNodestreams = xmlDoc.CreateElement("streams");
            root.AppendChild(childNodestreams);
           
            this.childNodestream = xmlDoc.CreateElement("stream");
            childNodestreams.AppendChild(childNodestream);
            childNodestream.SetAttribute("type", "XPATH");

            this.childNodeEntities = xmlDoc.CreateElement("entities");
            root.AppendChild(childNodeEntities);

            
            this.childNodeData = xmlDoc.CreateElement("datas");
            root.AppendChild(childNodeData);


            this.appEXCEL = new Microsoft.Office.Interop.Excel.ApplicationClass();
            // create the workbook object by opening  the excel file.
            workBook = appEXCEL.Workbooks.Open(excelFilePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            // Get The Active Worksheet Using Sheet Name Or Active Sheet
            workSheet = (Worksheet)workBook.ActiveSheet;
            this.excelRange = workSheet.Cells;

            childNodestream.SetAttribute("name", workSheet.Name);

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

                XmlElement childNode = xmlDoc.CreateElement("data");
                childNode.SetAttribute("name", excelLine.dataName);
                childNode.SetAttribute("type", excelLine.dataType);
                childNode.SetAttribute("length", excelLine.dataLong);
                XmlElement childNode2 = xmlDoc.CreateElement("description");
                XmlText textNode = xmlDoc.CreateTextNode("description");
                textNode.Value = excelLine.dataDescription;
                childNode2.AppendChild(textNode);
                childNode.AppendChild(childNode2);

                childNodeData.AppendChild(childNode);
            }
        }

        public override void createStream(BDOCExcelLine excelLine)
        {
            //STREAMS
            //first entity who defines the document separator
            if (excelLine.dataLevel == "1")
            {
                XmlElement streamStartNode = xmlDoc.CreateElement("separator");
                streamStartNode.SetAttribute("useForChildren", "true");
                streamStartNode.SetAttribute("type", "START");

                XmlElement streamStartNode2 = xmlDoc.CreateElement("xpath");
                XmlText textNode = xmlDoc.CreateTextNode("xpath");
                //If the excel has an specific XPATH value
                if (excelLine.dataXpath != "")
                    textNode.Value = "descendant-or-self::" + excelLine.dataXpath;
                else
                    textNode.Value = "descendant-or-self::" + excelLine.dataName;
                streamStartNode2.AppendChild(textNode);
                streamStartNode.AppendChild(streamStartNode2);
                childNodestream.AppendChild(streamStartNode);
            }

            // entity struct in excel file
            if (excelLine.dataTypeBDOC == "Entité" && excelLine.dataLevel != "0" && excelLine.dataLevel != "1")
            {
                streamNode = xmlDoc.CreateElement("entityNode");
                streamNode.SetAttribute("name", excelLine.dataName);
                streamNode.SetAttribute("iterative", (excelLine.dataIterative == "Non") ? "false" : "true");

                

                XmlElement streamNode2 = xmlDoc.CreateElement("xpath");
                XmlText textNode = xmlDoc.CreateTextNode("xpath");
                //If the excel has an specific XPATH value
                if (excelLine.dataXpath != "")
                    textNode.Value = "descendant-or-self::" + excelLine.dataXpath;
                else
                    textNode.Value = "descendant-or-self::" + excelLine.dataName;
                streamNode2.AppendChild(textNode);
                streamNode.AppendChild(streamNode2);

                childNodestream.AppendChild(streamNode);

                //pour les listes
                if (excelLine.dataName.IndexOf("LIST_") == 0)
                {
                    XmlElement dataStreamNode = xmlDoc.CreateElement("dataNode");
                    dataStreamNode.SetAttribute("name", excelLine.dataName);

                    XmlElement childNode2 = xmlDoc.CreateElement("xpath");
                    textNode = xmlDoc.CreateTextNode(excelLine.dataDescription);
                    //If the excel has an specific XPATH value
                    if (excelLine.dataXpath != "")
                        textNode.Value = "descendant-or-self::" + excelLine.dataXpath;
                    else
                    {
                        if (excelLine.dataName.Length > 18)
                            textNode.Value = "descendant-or-self::OCCURRENCES_" + excelLine.dataName.Substring(0, 18);
                        else
                            textNode.Value = "descendant-or-self::OCCURRENCES_" + excelLine.dataName;
                    }

                    childNode2.AppendChild(textNode);
                    dataStreamNode.AppendChild(childNode2);
                    streamNode.AppendChild(dataStreamNode);
                }
            }
            //Data struct in excel file
            if (excelLine.dataTypeBDOC != "Entité")
            {
                XmlElement dataStreamNode = xmlDoc.CreateElement("dataNode");
                dataStreamNode.SetAttribute("name", excelLine.dataName);
                XmlElement childNode2 = xmlDoc.CreateElement("xpath");
                XmlText textNode = xmlDoc.CreateTextNode(excelLine.dataDescription);
                if (excelLine.dataXpath != "")
                    textNode.Value = "descendant-or-self::" + excelLine.dataXpath;
                else
                    textNode.Value = "descendant-or-self::" + excelLine.dataName;
                childNode2.AppendChild(textNode);
                dataStreamNode.AppendChild(childNode2);

                if (streamNode != null)
                    streamNode.AppendChild(dataStreamNode);
                else
                {
                    throw new Exception(ConfigurationSettings.AppSettings["notEntity1Exception"] + excelLine.dataName + ConfigurationSettings.AppSettings["notEntity2Exception"]);
                }

            }
            //---------------------------------------

        }

        public override void createEntity(BDOCExcelLine excelLine)
        {
            //pour les entities           
            if (excelLine.dataTypeBDOC == "Entité" && excelLine.dataName != "FLUX" && excelLine.dataLevel != "0")
            {
                if (this.lastObjectInserted == "ENTITY")
                    childNodeEntities.RemoveChild(entityNode);

                entityNode = this.xmlDoc.CreateElement("entity");
                entityNode.SetAttribute("name", excelLine.dataName);
                entityNode.SetAttribute("iterative", (excelLine.dataIterative == "Non") ? "false" : "true");
                childNodeEntities.AppendChild(entityNode);

                //description de l entité
                XmlElement streamNodeDesc = xmlDoc.CreateElement("description");
                XmlText textNodeDesc = xmlDoc.CreateTextNode("description");
                textNodeDesc.Value = excelLine.dataDescription;
                streamNodeDesc.AppendChild(textNodeDesc);
                entityNode.AppendChild(streamNodeDesc);
                
                this.lastObjectInserted = "ENTITY";
                if (excelLine.dataName.IndexOf("LIST_") == 0)
                {
                    XmlElement dataEntityNode = xmlDoc.CreateElement("entityData");
                    XmlElement childNode2 = xmlDoc.CreateElement("description");
                    XmlText textNode = xmlDoc.CreateTextNode(excelLine.dataName);
                    if (excelLine.dataName.Length > 18)
                        textNode.Value = "OCCURRENCES_" + excelLine.dataName.Substring(0, 18);
                    else
                        textNode.Value = "OCCURRENCES_" + excelLine.dataName;

                    dataEntityNode.AppendChild(textNode);
                    entityNode.AppendChild(dataEntityNode);
                }
            }
            if (excelLine.dataTypeBDOC != "Entité")
            {
                XmlElement dataEntityNode = xmlDoc.CreateElement("entityData");
                XmlElement childNode2 = xmlDoc.CreateElement("description");
                XmlText textNode = xmlDoc.CreateTextNode(excelLine.dataDescription);
                textNode.Value = excelLine.dataName;
                dataEntityNode.AppendChild(textNode);
                entityNode.AppendChild(dataEntityNode);
                this.lastObjectInserted = "DATA";
            }
            //---------------------------------------

        }

        /// <summary>
        /// Fonction qui construira un XML à partir d'un Excel donné.
        /// </summary>
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
            index = 0;

            try
            {
                while (((Range)workSheet.Cells[rowIndex, colIndex1]).Value2 != null)
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

                this.FinalXML = this.xmlDoc.InnerXml;
                this.FinalTXT = this.FinalTXTData + "\r\n" + this.FinalTXTEntity;
                this.appEXCEL.Quit();
            }
            catch (Exception ex)
            {
                appEXCEL.Quit();
                Console.WriteLine(ex.Message + ", line " + index.ToString());
                throw ex;
            }
        }
    }
}
