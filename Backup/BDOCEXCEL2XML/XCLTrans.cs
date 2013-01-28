using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;


namespace BDOCEXCEL2XML
{
    class XCLTrans
    {
        private XmlDocument xmlDoc ;
        private XmlElement childNodeData;
        private XmlElement childNodeEntities;
        private XmlElement childNodestream;

        private XmlElement entityNode ;
        private XmlElement streamNode ;

        private Microsoft.Office.Interop.Excel.ApplicationClass appEXCEL;
        private Worksheet workSheet;
        private Workbook workBook;
        
        private string FinalXML;

        private string FinalTXT;
        private string FinalTXTData;
        private string FinalTXTEntity;
        private string FinalTXTStream;

        private string lastObjectInserted;
        public string XML
        {
            get
            {
                return this.FinalXML;
            }           
        }

        public string text
        {
            get
            {
                return this.FinalTXT;
            }
        }

        private BDOCExcelLine myBDOCExcelLine;
        
        public XCLTrans(string excelFilePath)
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
            childNodestream.SetAttribute("name", workSheet.Name);

            //************ TXT init
            FinalTXTData = "";
            FinalTXTEntity = "";
            FinalTXTStream = "";

           // throw new System.NotImplementedException();
        }

        private void createDataTXT(BDOCExcelLine excelLine)
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

        private void createData(BDOCExcelLine excelLine)
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

        private void createStream(BDOCExcelLine excelLine)
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
            if (excelLine.dataTypeBDOC != "Entité" )
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
                streamNode.AppendChild(dataStreamNode);
            }
            //---------------------------------------
           
        }

        private void createStreamTXT(BDOCExcelLine excelLine)
        {
            
            //pour les streamses
            //premiere entite qui definira le separateur de document
            if (excelLine.dataLevel == "1")
            {
                FinalTXTStream += "*" + "\r\n";
                FinalTXTStream += excelLine.dataName + "\r\n";
                FinalTXTStream += excelLine.dataDescription + "\r\n";
                FinalTXTStream += "0" + "\r\n"; //Stream type , only 0 available
                FinalTXTStream += "0" + "\r\n"; //Line type , only 0 available
                FinalTXTStream += "+0 +3" + "\r\n"; //TypeID ,see admin doc, page 69
                FinalTXTStream += "flotFILE.XML" + "\r\n"; //stream file path, 
                FinalTXTStream += "0" + "\r\n"; //Reserved , only 0 available
            }

            if (excelLine.dataTypeBDOC == "Entité" && excelLine.dataLevel != "0")
            {
                FinalTXTStream += "*E" + "\r\n";
                FinalTXTStream += excelLine.dataName + "\r\n";
                FinalTXTStream += excelLine.dataDescription + "\r\n";
                FinalTXTStream += "0" + "\r\n"; //Stream type , only 0 available
                FinalTXTStream += (excelLine.dataIterative == "Non") ? "0" : "1";                
                FinalTXTStream += "AAA" + "\r\n"; //Id, CORRECT IT!

                FinalTXTStream += "+0 +3" + "\r\n"; //TypeID ,see admin doc, page 69
                FinalTXTStream += "flotFILE.XML" + "\r\n"; //stream file path, 
                FinalTXTStream += "0" + "\r\n"; //Reserved , only 0 available
                streamNode.SetAttribute("iterative", (excelLine.dataIterative == "Non") ? "false" : "true");

                XmlElement streamNode2 = xmlDoc.CreateElement("xpath");
                XmlText textNode = xmlDoc.CreateTextNode("xpath");
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
                    if (excelLine.dataName.Length > 18)
                        textNode.Value = "descendant-or-self::OCCURRENCES_" + excelLine.dataName.Substring(0, 18);
                    else
                        textNode.Value = "descendant-or-self::OCCURRENCES_" + excelLine.dataName;
                    childNode2.AppendChild(textNode);
                    dataStreamNode.AppendChild(childNode2);
                    streamNode.AppendChild(dataStreamNode);
                }
            }
            if (excelLine.dataTypeBDOC != "Entité")
            {
                XmlElement dataStreamNode = xmlDoc.CreateElement("dataNode");
                dataStreamNode.SetAttribute("name", excelLine.dataName);
                XmlElement childNode2 = xmlDoc.CreateElement("xpath");
                XmlText textNode = xmlDoc.CreateTextNode(excelLine.dataDescription);
                textNode.Value = "descendant-or-self::" + excelLine.dataName;
                childNode2.AppendChild(textNode);
                dataStreamNode.AppendChild(childNode2);
                streamNode.AppendChild(dataStreamNode);
            }
            //---------------------------------------

        }

        private void createEntity(BDOCExcelLine excelLine)
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

        private void createEntityTXT(BDOCExcelLine excelLine)
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
                this.lastObjectInserted = "DATA" ;
            }
            //---------------------------------------

        }


        private void createStreamXMLFlux(BDOCExcelLine excelLine)
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
                streamNode.AppendChild(dataStreamNode);
            }
            //---------------------------------------

        }



        /// <summary>
        /// Fonction qui construira un XML à partir d'un Excel donné.
        /// </summary>
        public void createXML(ToolStripProgressBar MyProgress)
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
                    if (((Range)workSheet.Cells[rowIndex, colIndex1]).Value2 != null)
                    {
                        this.fillBDOCData(myBDOCExcelLine, rowIndex);

                        //XML
                        this.createData(myBDOCExcelLine);                       
                        this.createEntity(myBDOCExcelLine);
                        this.createStream(myBDOCExcelLine);

                        //TXT
                        //this.createStreamTXT(myBDOCExcelLine);
                        //this.createDataTXT(myBDOCExcelLine);
                        //this.createEntityTXT(myBDOCExcelLine);
                                                

                        MyProgress.PerformStep();

                        index++;                        
                    }
                }
                this.FinalXML = this.xmlDoc.InnerXml;

                this.FinalTXT = this.FinalTXTData + "\r\n" + this.FinalTXTEntity ;

                this.appEXCEL.Quit();        

            }
            catch (Exception ex)
            {
                appEXCEL.Quit();
                Console.WriteLine(ex.Message);
                throw ex;
            }    
            

        }

        public void createXMLFlux(ToolStripProgressBar MyProgress)
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
                    if (((Range)workSheet.Cells[rowIndex, colIndex1]).Value2 != null)
                    {
                        this.fillBDOCData(myBDOCExcelLine, rowIndex);

                        //XML                      
                        this.createStreamXMLFlux(myBDOCExcelLine);

                        MyProgress.PerformStep();

                        index++;
                    }
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

        public void fillBDOCData(BDOCExcelLine excelLine, Object rowIndex)
        {

            // This row,column index should be changed as per your need.
            // that is which cell in the excel you are interesting to read.
            object colIndex1 = 1;
            object colIndex2 = 2;
            object colIndex3 = 3;
            object colIndex4 = 4;
            object colIndex5 = 5;
            object colIndex6 = 6;
            object colIndex7 = 7;
            object colIndex8 = 8;
            object colIndex9 = 9;
            
            //read fields
            try
            {
                excelLine.reset();
                excelLine.dataLevel = (((Range)workSheet.Cells[rowIndex, colIndex1]).Value2 != null) ? ((Range)workSheet.Cells[rowIndex, colIndex1]).Value2.ToString().Trim() : "";
                excelLine.dataTypeBDOC = (((Range)workSheet.Cells[rowIndex, colIndex2]).Value2 != null) ? ((Range)workSheet.Cells[rowIndex, colIndex2]).Value2.ToString().Trim(): "";
                excelLine.dataName = (((Range)workSheet.Cells[rowIndex, colIndex4]).Value2 != null) ? ((Range)workSheet.Cells[rowIndex, colIndex4]).Value2.ToString().Trim(): "";
                excelLine.dataDescription = (((Range)workSheet.Cells[rowIndex, colIndex5]).Value2 != null) ? ((Range)workSheet.Cells[rowIndex, colIndex5]).Value2.ToString() : "";
                excelLine.dataXpath = (((Range)workSheet.Cells[rowIndex, colIndex8]).Value2 != null) ? ((Range)workSheet.Cells[rowIndex, colIndex8]).Value2.ToString() : "";
                excelLine.dataXMLFlux = (((Range)workSheet.Cells[rowIndex, colIndex9]).Value2 != null) ? ((Range)workSheet.Cells[rowIndex, colIndex9]).Value2.ToString() : "";

                if (excelLine.dataTypeBDOC == "Entité")
                {
                    excelLine.dataIterative = (((Range)workSheet.Cells[rowIndex, colIndex3]).Value2 != null) ? ((Range)workSheet.Cells[rowIndex, colIndex3]).Value2.ToString() : "";
                }
                else
                {
                    excelLine.dataType = (((Range)workSheet.Cells[rowIndex, colIndex6]).Value2 != null) ? ((Range)workSheet.Cells[rowIndex, colIndex6]).Value2.ToString() : "";
                    excelLine.dataLong = (((Range)workSheet.Cells[rowIndex, colIndex7]).Value2 != null) ? ((Range)workSheet.Cells[rowIndex, colIndex7]).Value2.ToString() : "";                    
                    excelLine.dataType = (excelLine.dataType == "INT") ? "INTEGER" : excelLine.dataType.ToUpper();

                }
            }
            catch (Exception ex) {
                throw new Exception("The values are not well formatted for the " + excelLine.dataName + " Field. " + ex.Message); 
            }

            //HERE ADD CheckBox FUNCTIONS for THE VALUES
         }
    }
}
