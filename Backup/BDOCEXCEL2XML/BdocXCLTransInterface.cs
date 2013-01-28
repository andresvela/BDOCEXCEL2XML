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
    public abstract class BdocXCLTransInterface
    {
        public XmlDocument xmlDoc;
        public XmlElement childNodeData;
        public XmlElement childNodeEntities;
        public XmlElement childNodestream;

        public XmlElement entityNode;
        public XmlElement streamNode;

        public Microsoft.Office.Interop.Excel.ApplicationClass appEXCEL;
        public Worksheet workSheet;
        public Workbook workBook;
        public Range excelRange;

        public string FinalXML;

        public string FinalTXT;
        public string FinalTXTData;
        public string FinalTXTEntity;
        public string FinalTXTStream;

        public int total; // registers total 

        public string lastObjectInserted;


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

        public BDOCExcelLine myBDOCExcelLine;

        public BdocXCLTransInterface()
        {
            
        }

       
        public virtual void createData(BDOCExcelLine excelLine)
        {
           
        }

        public virtual void createStream(BDOCExcelLine excelLine)
        {           
           
        }

        public virtual void createEntity(BDOCExcelLine excelLine)
        {            
                       
        }
        
        /// <summary>
        /// Fonction qui construira un XML à partir d'un Excel donné.
        /// </summary>
        public virtual void createXML(ToolStripProgressBar MyProgress)
        {
                   
        }

        public void fillBDOCData(BDOCExcelLine excelLine, Object rowIndex)
        {

            // This row,column index should be changed as per your need.
            // that is which cell in the excel you are interesting to read.
            int colIndex1 = 1;
            int colIndex2 = 2;
            int colIndex3 = 3;
            int colIndex4 = 4;
            int colIndex5 = 5;
            int colIndex6 = 6;
            int colIndex7 = 7;
            int colIndex8 = 8;
            int colIndex9 = 9;

            Object dataLevel = ((Range)excelRange[rowIndex, colIndex1]).Value2;
            Object dataTypeBDOC = ((Range)excelRange[rowIndex, colIndex2]).Value2;
            Object dataName = ((Range)excelRange[rowIndex, colIndex4]).Value2;
            Object dataDescription = ((Range)excelRange[rowIndex, colIndex5]).Value2;
            Object dataXpath = ((Range)excelRange[rowIndex, colIndex8]).Value2;
            Object dataXMLFlux = ((Range)excelRange[rowIndex, colIndex9]).Value2;
            
            if (dataLevel == null)
            {
                return;
            }

            //read fields
            try
            {
                excelLine.reset();
                                
                excelLine.dataLevel = (dataLevel != null) ? dataLevel.ToString().Trim() : "";
                excelLine.dataTypeBDOC = ( dataTypeBDOC != null) ? dataTypeBDOC.ToString().Trim() : "";
                excelLine.dataName = (dataName != null) ? dataName.ToString().Trim() : "";
                excelLine.dataDescription = (dataDescription != null) ? dataDescription.ToString() : "";
                excelLine.dataXpath = (dataXpath != null) ? dataXpath.ToString() : "";
                excelLine.dataXMLFlux = (dataXMLFlux != null) ? dataXMLFlux.ToString() : "";

                if (excelLine.dataTypeBDOC == "Entité")
                {
                    Object dataIterative = ((Range)excelRange[rowIndex, colIndex3]).Value2;
                    excelLine.dataIterative = (dataIterative != null) ? dataIterative.ToString() : "";
                }
                else
                {
                    Object dataType = ((Range)excelRange[rowIndex, colIndex6]).Value2;
                    Object dataLong = ((Range)excelRange[rowIndex, colIndex7]).Value2;
                                        
                    excelLine.dataType = (dataType != null) ? dataType.ToString() : "";
                    excelLine.dataLong = (dataLong!= null) ? dataLong.ToString() : "";                    
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
