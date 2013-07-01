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
    class BDOCXCLTransImportV5 : BdocXCLTransInterface
    {
        public BDOCXCLTransImportV5(string excelFilePath, string type)
        {
             this.myBDOCExcelLine = new BDOCExcelLine();

            //XML ROOT node
            this.xmlDoc = new XmlDocument();
            //XmlDeclaration dec = this.xmlDoc.CreateXmlDeclaration("1.0", "ISO-8859-1", null);
            XmlDeclaration dec = this.xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
            this.xmlDoc.AppendChild(dec);
            XmlElement rootNode = xmlDoc.CreateElement("PackageDTO");
            rootNode.SetAttribute("xmlns:xsd", "http://www.bdoc.com");
            rootNode.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema");                        
            this.xmlDoc.AppendChild(rootNode);

            //all the resting structure
            XmlNode root = xmlDoc.DocumentElement;
            
            //XML TRANSALTION CODE LISTS
            XmlElement TranslationCodesList = xmlDoc.CreateElement("TranslationCodesList");
            root.AppendChild(TranslationCodesList);

            //XmlElement TranslationCodeDTO = xmlDoc.CreateElement("TranslationCodeDTO");
            //TranslationCodesList.AppendChild(TranslationCodeDTO);

            //XmlElement TranslationCodeName = xmlDoc.CreateElement("Name");
            //XmlElement TranslationCodeDescription = xmlDoc.CreateElement("Description");

            //TranslationCodeDTO.AppendChild(TranslationCodeName);
            //TranslationCodeDTO.AppendChild(TranslationCodeDescription);


            //XML DATA LISTS
            this.childNodeData = xmlDoc.CreateElement("DataList");
            root.AppendChild(childNodeData);

            //XML ENTITIIES LIST
            childNodeEntities = xmlDoc.CreateElement("EntitiesList");
            root.AppendChild(childNodeEntities);

            //streams
            XmlElement DatastreamsList = xmlDoc.CreateElement("DatastreamsList");
            root.AppendChild(DatastreamsList);

            XmlElement DatastreamDTO = xmlDoc.CreateElement("DatastreamDTO");
            DatastreamsList.AppendChild(DatastreamDTO);

            
           
            this.appEXCEL = new Microsoft.Office.Interop.Excel.ApplicationClass();
            // create the workbook object by opening  the excel file.
            workBook = appEXCEL.Workbooks.Open(excelFilePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            // Get The Active Worksheet Using Sheet Name Or Active Sheet
            workSheet = (Worksheet)workBook.ActiveSheet;
            this.excelRange = workSheet.Cells;

            XmlElement NameNodestream = xmlDoc.CreateElement("Name");
            XmlText nameNodeText = xmlDoc.CreateTextNode("Name");
            nameNodeText.Value = workSheet.Name.ToUpper();
            NameNodestream.AppendChild(nameNodeText);
            DatastreamDTO.AppendChild(NameNodestream);

            XmlElement TypeNodestream = xmlDoc.CreateElement("DatastreamType");
            XmlText typeNodeText = xmlDoc.CreateTextNode("DatastreamType");

            if (type == "XML")
            {
                typeNodeText.Value = "XPATH";
            }
            if (type == "FLAT")
            {
                typeNodeText.Value = "FLAT";
            }
                    
           
           
            TypeNodestream.AppendChild(typeNodeText);
            DatastreamDTO.AppendChild(TypeNodestream);


            if (type == "XML")
            {
                this.childNodestream = xmlDoc.CreateElement("XPathDatastream");
            }
            if (type == "FLAT")
            {
                this.childNodestream = xmlDoc.CreateElement("FlatDatastream");
                XmlElement FlatEncodingNodestream = xmlDoc.CreateElement("FlatEncoding");
                XmlText FlatEncodingNodeText = xmlDoc.CreateTextNode("FlatEncoding");
                FlatEncodingNodeText.Value = "Win1252";
                FlatEncodingNodestream.AppendChild(FlatEncodingNodeText);
                this.childNodestream.AppendChild(FlatEncodingNodestream);
            }
          
           
            DatastreamDTO.AppendChild(childNodestream);

           

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

                XmlElement childNode = xmlDoc.CreateElement("DataDTO");

                XmlElement nameNode = xmlDoc.CreateElement("Name");
                XmlText nameNodeText = xmlDoc.CreateTextNode("Name");
                nameNodeText.Value = excelLine.dataName;
                nameNode.AppendChild(nameNodeText);

                XmlElement descriptionNode = xmlDoc.CreateElement("Description");
                XmlText descriptionNodeText = xmlDoc.CreateTextNode("Description");
                descriptionNodeText.Value = excelLine.dataDescription;
                descriptionNode.AppendChild(descriptionNodeText);

                XmlElement typeNode = xmlDoc.CreateElement("Type");
                XmlText typeNodeText = xmlDoc.CreateTextNode("Type");
                typeNodeText.Value = excelLine.dataType;
                typeNode.AppendChild(typeNodeText);

                XmlElement DescriptionTranslationsList = xmlDoc.CreateElement("DescriptionTranslationsList");
                //XmlElement DescriptionTranslationDTO = xmlDoc.CreateElement("DescriptionTranslationDTO");

                //XmlElement TranslationCodeName = xmlDoc.CreateElement("TranslationCodeName");
                //XmlElement TranslationCodeDescription = xmlDoc.CreateElement("Translation");

                //DescriptionTranslationDTO.AppendChild(TranslationCodeName);
                //DescriptionTranslationDTO.AppendChild(TranslationCodeDescription);


                childNode.AppendChild(nameNode);
                childNode.AppendChild(descriptionNode);
                childNode.AppendChild(typeNode);

                //DescriptionTranslationsList.AppendChild(DescriptionTranslationDTO);

                childNode.AppendChild(DescriptionTranslationsList);

                childNodeData.AppendChild(childNode);
            }
        }

        /*        public override void createStream(BDOCExcelLine excelLine)
                {
                    //STREAMS
                    //first entity who defines the document separator
                    if (excelLine.dataLevel == "1")
                    {
                        XmlElement streamStartNode = xmlDoc.CreateElement("XPathRequestStartDoc");

                        XmlText textNode = xmlDoc.CreateTextNode("XPathRequestStartDoc");
                        //If the excel has an specific XPATH value
                        if (excelLine.dataXpath != "")
                            textNode.Value = "descendant-or-self::" + excelLine.dataXpath;
                        else
                            textNode.Value = "descendant-or-self::" + excelLine.dataName;
                        streamStartNode.AppendChild(textNode);
                        childNodestream.AppendChild(streamStartNode);
                        this.XPathEntitiesList = xmlDoc.CreateElement("XPathEntitiesList");
                        childNodestream.AppendChild(XPathEntitiesList);

                    }

                    // entity struct in excel file
                    if (excelLine.dataTypeBDOC == "Entité" && excelLine.dataLevel != "0" && excelLine.dataLevel != "1")
                    {
                        streamNode = xmlDoc.CreateElement("XPathEntityDTO");
                        XmlElement RefEntityName = xmlDoc.CreateElement("RefEntityName");
                        XmlText RefEntityNameText = xmlDoc.CreateTextNode("RefEntityName");
                        RefEntityNameText.Value =  excelLine.dataName;

                        this.XPathDataList = xmlDoc.CreateElement("XPathDataList");

                        XmlElement streamNode2 = xmlDoc.CreateElement("XpathRequest");
                        XmlText streamNode2Text = xmlDoc.CreateTextNode("XpathRequest");
                        //If the excel has an specific XPATH value
                        if (excelLine.dataXpath != "")
                            streamNode2Text.Value = "descendant-or-self::" + excelLine.dataXpath;
                        else
                            streamNode2Text.Value = "descendant-or-self::" + excelLine.dataName;

                        streamNode2.AppendChild(streamNode2Text);
                        streamNode.AppendChild(streamNode2);
                        RefEntityName.AppendChild(RefEntityNameText);
                        streamNode.AppendChild(RefEntityName);
                        streamNode.AppendChild(this.XPathDataList);

                        XPathEntitiesList.AppendChild(streamNode);

                        //pour les listes
                        if (excelLine.dataName.IndexOf("LIST_") == 0)
                        {
                            XmlElement dataStreamNode = xmlDoc.CreateElement("dataNode");
                            dataStreamNode.SetAttribute("name", excelLine.dataName);

                            XmlElement childNode2 = xmlDoc.CreateElement("xpath");
                            streamNode2Text = xmlDoc.CreateTextNode(excelLine.dataDescription);
                            //If the excel has an specific XPATH value
                            if (excelLine.dataXpath != "")
                                streamNode2Text.Value = "descendant-or-self::" + excelLine.dataXpath;
                            else
                            {
                                if (excelLine.dataName.Length > 18)
                                    streamNode2Text.Value = "descendant-or-self::OCCURRENCES_" + excelLine.dataName.Substring(0, 18);
                                else
                                    streamNode2Text.Value = "descendant-or-self::OCCURRENCES_" + excelLine.dataName;
                            }

                            childNode2.AppendChild(streamNode2Text);
                            dataStreamNode.AppendChild(childNode2);
                            streamNode.AppendChild(dataStreamNode);
                        }
                    }

                    //Data struct in excel file
                    if (excelLine.dataTypeBDOC != "Entité")
                    {
                        XmlElement dataStreamNode = xmlDoc.CreateElement("XPathDataDTO");

                        XmlElement RefDataName = xmlDoc.CreateElement("RefDataName");
                        XmlText RefDataNameText = xmlDoc.CreateTextNode("RefDataName");
                        RefDataNameText.Value = excelLine.dataName;
                        RefDataName.AppendChild(RefDataNameText);


                        XmlElement DataFormat = xmlDoc.CreateElement("DataFormat");


                        XmlElement NumericDataFormat = xmlDoc.CreateElement("NumericDataFormat");
                        XmlElement DecimalSeparator = xmlDoc.CreateElement("DecimalSeparator");
                        XmlText DecimalSeparatorVal = xmlDoc.CreateTextNode("DecimalSeparator");
                        DecimalSeparatorVal.Value = ",";
                        DecimalSeparator.AppendChild(DecimalSeparatorVal);
                        NumericDataFormat.AppendChild(DecimalSeparator);

                        XmlElement DatetimeDataFormat = xmlDoc.CreateElement("DatetimeDataFormat");
                        XmlElement DatetimeFormat = xmlDoc.CreateElement("DatetimeFormat");
                        XmlElement DateTimeFormatProduction = xmlDoc.CreateElement("DateTimeFormatProduction");

                        XmlText DatetimeFormatVal = xmlDoc.CreateTextNode("DatetimeFormat");
                        XmlText DateTimeFormatProductionVal = xmlDoc.CreateTextNode("DateTimeFormatProduction");
                        DatetimeFormatVal.Value = "dd/MM/yyyy";
                        DatetimeFormat.AppendChild(DatetimeFormatVal);
                        DateTimeFormatProductionVal.Value = "%d/%m/%Y";
                        DateTimeFormatProduction.AppendChild(DateTimeFormatProductionVal);

                        DatetimeDataFormat.AppendChild(DatetimeFormat);
                        DatetimeDataFormat.AppendChild(DateTimeFormatProduction);

                        XmlElement StringDataFormat = xmlDoc.CreateElement("StringDataFormat");
                

                        DataFormat.AppendChild(NumericDataFormat);
                        DataFormat.AppendChild(DatetimeDataFormat);
                        DataFormat.AppendChild(StringDataFormat);
                
                        XmlElement childNode2 = xmlDoc.CreateElement("XpathRequest");
                        XmlText textNode = xmlDoc.CreateTextNode(excelLine.dataDescription);
                        if (excelLine.dataXpath != "")
                            textNode.Value = "descendant-or-self::" + excelLine.dataXpath;
                        else
                            textNode.Value = "descendant-or-self::" + excelLine.dataName;
                        childNode2.AppendChild(textNode);
                        dataStreamNode.AppendChild(childNode2);
                        dataStreamNode.AppendChild(RefDataName);
                        dataStreamNode.AppendChild(DataFormat);

                        if (streamNode != null)
                            this.XPathDataList.AppendChild(dataStreamNode);
                        else
                        {
                            throw new Exception(ConfigurationSettings.AppSettings["notEntity1Exception"] + excelLine.dataName + ConfigurationSettings.AppSettings["notEntity2Exception"]);
                        }

                    }
                    //---------------------------------------

                }
*/
                public override void createEntity(BDOCExcelLine excelLine)
                {
                    //pour les entities           
                    if (excelLine.dataTypeBDOC == "Entité" && excelLine.dataName != "FLUX" && excelLine.dataLevel != "0")
                    {
                        if (this.lastObjectInserted == "ENTITY")
                            childNodeEntities.RemoveChild(entityNode);

            

                        entityNode = this.xmlDoc.CreateElement("EntityDTO");               
                        childNodeEntities.AppendChild(entityNode);


                        XmlElement nameNode = xmlDoc.CreateElement("Name");
                        XmlText nameNodeText = xmlDoc.CreateTextNode("Name");
                        nameNodeText.Value = excelLine.dataName;
                        nameNode.AppendChild(nameNodeText);

                        XmlElement descriptionNode = xmlDoc.CreateElement("Description");
                        XmlText descriptionNodeText = xmlDoc.CreateTextNode("Description");
                        descriptionNodeText.Value = excelLine.dataDescription;
                        descriptionNode.AppendChild(descriptionNodeText);

                        XmlElement typeNode = xmlDoc.CreateElement("IsIterative");
                        XmlText typeNodeText = xmlDoc.CreateTextNode("IsIterative");
                        typeNodeText.Value = (excelLine.dataIterative == "Non") ? "false" : "true";
                        typeNode.AppendChild(typeNodeText);


                        this.dataEntityListNode = xmlDoc.CreateElement("DataNamesList");
             

                        entityNode.AppendChild(nameNode);
                        entityNode.AppendChild(descriptionNode);
                        entityNode.AppendChild(typeNode);
                        entityNode.AppendChild(this.dataEntityListNode);

                          
                        this.lastObjectInserted = "ENTITY";
                        if (excelLine.dataName.IndexOf("LIST_") == 0)
                        {
                            XmlElement dataEntityNode = xmlDoc.CreateElement("string");                    
                            XmlText textNode = xmlDoc.CreateTextNode(excelLine.dataName);
                            if (excelLine.dataName.Length > 18)
                                textNode.Value = "OCCURRENCES_" + excelLine.dataName.Substring(0, 18);
                            else
                                textNode.Value = "OCCURRENCES_" + excelLine.dataName;

                            dataEntityNode.AppendChild(textNode);
                            this.dataEntityListNode.AppendChild(dataEntityNode);
                        }
                    }
                    if (excelLine.dataTypeBDOC != "Entité")
                    {
                        XmlElement dataEntityNode = xmlDoc.CreateElement("string");               
                        XmlText textNode = xmlDoc.CreateTextNode(excelLine.dataDescription);
                        textNode.Value = excelLine.dataName;
                        dataEntityNode.AppendChild(textNode);
                        this.dataEntityListNode.AppendChild(dataEntityNode);
                        this.lastObjectInserted = "DATA";
                    }
                    //---------------------------------------

                }

     
      
        public override void createStream(BDOCExcelLine excelLine)
        {
            //STREAMS
            //first entity who defines the document separator
            if (excelLine.dataLevel == "1")
            {
                XmlElement streamStartNode = xmlDoc.CreateElement("XPathRequestStartDoc");

                
                XmlText IdEntitytext = xmlDoc.CreateTextNode("IdEntity");
                //If the excel has an specific XPATH value               
                IdEntitytext.Value = excelLine.dataXpath;
                streamStartNode.AppendChild(IdEntitytext);
               

                childNodestream.AppendChild(streamStartNode);


                this.XPathEntitiesList = xmlDoc.CreateElement("XPathEntitiesList");
                childNodestream.AppendChild(XPathEntitiesList);

            }

            // entity struct in excel file
            if (excelLine.dataTypeBDOC == "Entité" && excelLine.dataLevel != "0" && excelLine.dataLevel != "1")
            {
                streamNode = xmlDoc.CreateElement("XPathEntityDTO");
                XmlElement RefEntityName = xmlDoc.CreateElement("RefEntityName");
                XmlText RefEntityNameText = xmlDoc.CreateTextNode("RefEntityName");
                RefEntityNameText.Value = excelLine.dataName;

                this.XPathDataList = xmlDoc.CreateElement("XPathDataList");

                XmlElement streamNode2 = xmlDoc.CreateElement("XpathRequest");
                XmlText streamNode2Text = xmlDoc.CreateTextNode("XpathRequest");
                //If the excel has an specific XPATH value
                if (excelLine.dataXpath != "")
                    streamNode2Text.Value =  excelLine.dataXpath;
                else
                    streamNode2Text.Value =  excelLine.dataName;

                streamNode2.AppendChild(streamNode2Text);
                streamNode.AppendChild(streamNode2);
                RefEntityName.AppendChild(RefEntityNameText);
                streamNode.AppendChild(RefEntityName);
                streamNode.AppendChild(this.XPathDataList);

                XPathEntitiesList.AppendChild(streamNode);

                //pour les listes
                if (excelLine.dataName.IndexOf("LIST_") == 0)
                {
                    XmlElement dataStreamNode = xmlDoc.CreateElement("dataNode");
                    dataStreamNode.SetAttribute("name", excelLine.dataName);

                    XmlElement childNode2 = xmlDoc.CreateElement("xpath");
                    streamNode2Text = xmlDoc.CreateTextNode(excelLine.dataDescription);
                    //If the excel has an specific XPATH value
                    if (excelLine.dataXpath != "")
                        streamNode2Text.Value = "descendant-or-self::" + excelLine.dataXpath;
                    else
                    {
                        if (excelLine.dataName.Length > 18)
                            streamNode2Text.Value = "descendant-or-self::OCCURRENCES_" + excelLine.dataName.Substring(0, 18);
                        else
                            streamNode2Text.Value = "descendant-or-self::OCCURRENCES_" + excelLine.dataName;
                    }

                    childNode2.AppendChild(streamNode2Text);
                    dataStreamNode.AppendChild(childNode2);
                    streamNode.AppendChild(dataStreamNode);
                }
            }

            //Data struct in excel file
            if (excelLine.dataTypeBDOC != "Entité")
            {
                XmlElement dataStreamNode = xmlDoc.CreateElement("XPathDataDTO");

                XmlElement RefDataName = xmlDoc.CreateElement("RefDataName");
                XmlText RefDataNameText = xmlDoc.CreateTextNode("RefDataName");
                RefDataNameText.Value = excelLine.dataName;
                RefDataName.AppendChild(RefDataNameText);


                XmlElement XPathDataDTO = xmlDoc.CreateElement("XpathRequest");
                XmlText XPathDataDTOText = xmlDoc.CreateTextNode("XpathRequest");
                XPathDataDTOText.Value = excelLine.dataXpath;
                XPathDataDTO.AppendChild(XPathDataDTOText);

                                
                XmlElement DataFormat = xmlDoc.CreateElement("DataFormat");


                XmlElement NumericDataFormat = xmlDoc.CreateElement("NumericDataFormat");
                XmlElement DecimalSeparator = xmlDoc.CreateElement("DecimalSeparator");
                XmlText DecimalSeparatorVal = xmlDoc.CreateTextNode("DecimalSeparator");
                DecimalSeparatorVal.Value = ",";
                DecimalSeparator.AppendChild(DecimalSeparatorVal);
                NumericDataFormat.AppendChild(DecimalSeparator);

                XmlElement DatetimeDataFormat = xmlDoc.CreateElement("DatetimeDataFormat");
                XmlElement DatetimeFormat = xmlDoc.CreateElement("DatetimeFormat");
                XmlElement DateTimeFormatProduction = xmlDoc.CreateElement("DateTimeFormatProduction");

                XmlText DatetimeFormatVal = xmlDoc.CreateTextNode("DatetimeFormat");
                XmlText DateTimeFormatProductionVal = xmlDoc.CreateTextNode("DateTimeFormatProduction");
                DatetimeFormatVal.Value = "dd/MM/yyyy";
                DatetimeFormat.AppendChild(DatetimeFormatVal);
                DateTimeFormatProductionVal.Value = "%d/%m/%Y";
                DateTimeFormatProduction.AppendChild(DateTimeFormatProductionVal);

                DatetimeDataFormat.AppendChild(DatetimeFormat);
                DatetimeDataFormat.AppendChild(DateTimeFormatProduction);

                XmlElement StringDataFormat = xmlDoc.CreateElement("StringDataFormat");


                DataFormat.AppendChild(NumericDataFormat);
                DataFormat.AppendChild(DatetimeDataFormat);
                DataFormat.AppendChild(StringDataFormat);

               



               // XmlElement childNode2 = xmlDoc.CreateElement("XpathRequest");
               // XmlText textNode = xmlDoc.CreateTextNode(excelLine.dataDescription);
               // if (excelLine.dataXpath != "")
                //    textNode.Value = "descendant-or-self::" + excelLine.dataXpath;
                //else
               //     textNode.Value = "descendant-or-self::" + excelLine.dataName;
               // childNode2.AppendChild(textNode);
                //dataStreamNode.AppendChild(childNode2);
                dataStreamNode.AppendChild(RefDataName);
                dataStreamNode.AppendChild(XPathDataDTO);               

                dataStreamNode.AppendChild(DataFormat);

                if (streamNode != null)
                    this.XPathDataList.AppendChild(dataStreamNode);
                else
                {
                    throw new Exception(ConfigurationSettings.AppSettings["notEntity1Exception"] + excelLine.dataName + ConfigurationSettings.AppSettings["notEntity2Exception"]);
                }

            }
            //---------------------------------------

        }


        public override void createStreamFlat(BDOCExcelLine excelLine)
        {
            //STREAMS
            //first entity who defines the document separator
            if (excelLine.dataLevel == "1")
            {
                XmlElement streamStartNode = xmlDoc.CreateElement("FlatStartDocEntity");

                XmlElement IdEntity = xmlDoc.CreateElement("IdEntity");
                XmlText IdEntitytext = xmlDoc.CreateTextNode("IdEntity");
                //If the excel has an specific XPATH value               
                IdEntitytext.Value = excelLine.dataXpath;
                IdEntity.AppendChild(IdEntitytext);
                streamStartNode.AppendChild(IdEntity);

                childNodestream.AppendChild(streamStartNode);

                XmlElement StartSeparator = xmlDoc.CreateElement("StartSeparator");
                XmlElement Position = xmlDoc.CreateElement("Position");
                XmlText PositionText = xmlDoc.CreateTextNode("Position");
                PositionText.Value = "2";
                Position.AppendChild(PositionText);
                StartSeparator.AppendChild(Position);

                XmlElement StopSeparator = xmlDoc.CreateElement("StopSeparator");
                XmlElement Value = xmlDoc.CreateElement("Value");
                XmlText ValueText = xmlDoc.CreateTextNode("Value");
                ValueText.Value = "#";
                Value.AppendChild(ValueText);
                StopSeparator.AppendChild(Value);

                streamStartNode.AppendChild(StartSeparator);
                streamStartNode.AppendChild(StopSeparator);

                this.XPathEntitiesList = xmlDoc.CreateElement("FlatEntitiesList");
                childNodestream.AppendChild(XPathEntitiesList);

            }

            // entity struct in excel file
            if (excelLine.dataTypeBDOC == "Entité" && excelLine.dataLevel != "0" && excelLine.dataLevel != "1")
            {
                streamNode = xmlDoc.CreateElement("FlatEntityDTO");
                XmlElement RefEntityName = xmlDoc.CreateElement("RefEntityName");
                XmlText RefEntityNameText = xmlDoc.CreateTextNode("RefEntityName");
                RefEntityNameText.Value = excelLine.dataName;

                this.XPathDataList = xmlDoc.CreateElement("FlatDataFillersList");

                XmlElement streamNode2 = xmlDoc.CreateElement("IdEntity");
                XmlText streamNode2Text = xmlDoc.CreateTextNode("IdEntity");
                //If the excel has an specific XPATH value
                if (excelLine.dataXpath != "")
                    streamNode2Text.Value = excelLine.dataXpath;
                else
                    streamNode2Text.Value = excelLine.dataName;

                streamNode2.AppendChild(streamNode2Text);
                streamNode.AppendChild(streamNode2);
                RefEntityName.AppendChild(RefEntityNameText);
                streamNode.AppendChild(RefEntityName);
                streamNode.AppendChild(this.XPathDataList);

                XPathEntitiesList.AppendChild(streamNode);

                //pour les listes
                if (excelLine.dataName.IndexOf("LIST_") == 0)
                {
                    XmlElement dataStreamNode = xmlDoc.CreateElement("dataNode");
                    dataStreamNode.SetAttribute("name", excelLine.dataName);

                    XmlElement childNode2 = xmlDoc.CreateElement("xpath");
                    streamNode2Text = xmlDoc.CreateTextNode(excelLine.dataDescription);
                    //If the excel has an specific XPATH value
                    if (excelLine.dataXpath != "")
                        streamNode2Text.Value = "descendant-or-self::" + excelLine.dataXpath;
                    else
                    {
                        if (excelLine.dataName.Length > 18)
                            streamNode2Text.Value = "descendant-or-self::OCCURRENCES_" + excelLine.dataName.Substring(0, 18);
                        else
                            streamNode2Text.Value = "descendant-or-self::OCCURRENCES_" + excelLine.dataName;
                    }

                    childNode2.AppendChild(streamNode2Text);
                    dataStreamNode.AppendChild(childNode2);
                    streamNode.AppendChild(dataStreamNode);
                }
            }

            //Data struct in excel file
            if (excelLine.dataTypeBDOC != "Entité")
            {
                XmlElement dataStreamNode = xmlDoc.CreateElement("FlatDataFillerDTO");

                XmlElement RefDataName = xmlDoc.CreateElement("RefDataName");
                XmlText RefDataNameText = xmlDoc.CreateTextNode("RefDataName");
                RefDataNameText.Value = excelLine.dataName;
                RefDataName.AppendChild(RefDataNameText);


                XmlElement StartSeparator = xmlDoc.CreateElement("StartSeparator");
                XmlElement Position = xmlDoc.CreateElement("Position");
                XmlText PositionText = xmlDoc.CreateTextNode("Position");
                PositionText.Value = "1";
                Position.AppendChild(PositionText);
                StartSeparator.AppendChild(Position);

                XmlElement StopSeparator = xmlDoc.CreateElement("StopSeparator");
                XmlElement Value = xmlDoc.CreateElement("Value");
                XmlText ValueText = xmlDoc.CreateTextNode("Value");
                ValueText.Value = "#";
                Value.AppendChild(ValueText);
                StopSeparator.AppendChild(Value);


                XmlElement DataFormat = xmlDoc.CreateElement("DataFormat");


                XmlElement NumericDataFormat = xmlDoc.CreateElement("NumericDataFormat");
                XmlElement DecimalSeparator = xmlDoc.CreateElement("DecimalSeparator");
                XmlText DecimalSeparatorVal = xmlDoc.CreateTextNode("DecimalSeparator");
                DecimalSeparatorVal.Value = ",";
                DecimalSeparator.AppendChild(DecimalSeparatorVal);
                NumericDataFormat.AppendChild(DecimalSeparator);

                XmlElement DatetimeDataFormat = xmlDoc.CreateElement("DatetimeDataFormat");
                XmlElement DatetimeFormat = xmlDoc.CreateElement("DatetimeFormat");
                XmlElement DateTimeFormatProduction = xmlDoc.CreateElement("DateTimeFormatProduction");

                XmlText DatetimeFormatVal = xmlDoc.CreateTextNode("DatetimeFormat");
                XmlText DateTimeFormatProductionVal = xmlDoc.CreateTextNode("DateTimeFormatProduction");
                DatetimeFormatVal.Value = "dd/MM/yyyy";
                DatetimeFormat.AppendChild(DatetimeFormatVal);
                DateTimeFormatProductionVal.Value = "%d/%m/%Y";
                DateTimeFormatProduction.AppendChild(DateTimeFormatProductionVal);

                DatetimeDataFormat.AppendChild(DatetimeFormat);
                DatetimeDataFormat.AppendChild(DateTimeFormatProduction);

                XmlElement StringDataFormat = xmlDoc.CreateElement("StringDataFormat");


                DataFormat.AppendChild(NumericDataFormat);
                DataFormat.AppendChild(DatetimeDataFormat);
                DataFormat.AppendChild(StringDataFormat);





                // XmlElement childNode2 = xmlDoc.CreateElement("XpathRequest");
                // XmlText textNode = xmlDoc.CreateTextNode(excelLine.dataDescription);
                // if (excelLine.dataXpath != "")
                //    textNode.Value = "descendant-or-self::" + excelLine.dataXpath;
                //else
                //     textNode.Value = "descendant-or-self::" + excelLine.dataName;
                // childNode2.AppendChild(textNode);
                //dataStreamNode.AppendChild(childNode2);
                dataStreamNode.AppendChild(RefDataName);
                dataStreamNode.AppendChild(StartSeparator);
                dataStreamNode.AppendChild(StopSeparator);

                dataStreamNode.AppendChild(DataFormat);

                if (streamNode != null)
                    this.XPathDataList.AppendChild(dataStreamNode);
                else
                {
                    throw new Exception(ConfigurationSettings.AppSettings["notEntity1Exception"] + excelLine.dataName + ConfigurationSettings.AppSettings["notEntity2Exception"]);
                }

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
                    
                    if (type == "XML") {
                        this.createStream(myBDOCExcelLine);
                    }
                    if (type=="FLAT"){
                        this.createStreamFlat(myBDOCExcelLine);
                    }
                    

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
