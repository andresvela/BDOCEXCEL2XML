using System;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Xml;
using System.IO;

namespace BDOCEXCEL2XML
{
    public partial class Form1 : Form
    {
        public Workbook workBook;
        public Form1()
        {
            InitializeComponent();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            txtB_XCLFile.Text = XCLFileDialog.FileName;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void btn_XCLOpen_Click(object sender, EventArgs e)
        {
            XCLFileDialog.ShowDialog();
            if (XCLFileDialog.FileName != "")
            {
                Microsoft.Office.Interop.Excel.ApplicationClass appEXCEL = new Microsoft.Office.Interop.Excel.ApplicationClass();
                // create the workbook object by opening  the excel file.
                //Workbook workBook = appEXCEL.Workbooks.Open(XCLFileDialog.FileName, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                this.workBook = appEXCEL.Workbooks.Open(XCLFileDialog.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Get The Active Worksheet Using Sheet Name Or Active Sheet



                foreach (Worksheet mySheet in workBook.Sheets)
                {
                    //Worksheet mySheet = (Worksheet)workBook.Sheets.get_Item(i);
                    comboBox2.Items.Add(mySheet.Name);
                }
                                
                Worksheet workSheet = (Worksheet)workBook.ActiveSheet;
                comboBox2.SelectedItem = workSheet.Name;

                lbl_messages.Text = ConfigurationSettings.AppSettings["sheetText1"] + workSheet.Name + ConfigurationSettings.AppSettings["sheetText2"];
                toolStripStatusLabel2.Text = ConfigurationSettings.AppSettings["changeSheetText"]; 
                appEXCEL.Quit();
                button1.Enabled = true;
                button2.Enabled = false;
                btCreateXMLFlux.Enabled = true;
                button3.Enabled = true;

              
                
                webBrowser1.DocumentText = ConfigurationSettings.AppSettings["excelReadyText"];
                webBrowser2.DocumentText = ConfigurationSettings.AppSettings["excelReadyText"];
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel2.Text = ConfigurationSettings.AppSettings["workingText"]; 

            BdocXCLTransImport myXCLTrans = new BdocXCLTransImport(XCLFileDialog.FileName);
            try
            {
                myXCLTrans.createXML(toolStripProgressBar1);
                webBrowser1.DocumentText = myXCLTrans.XML ;
                
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.InnerXml = myXCLTrans.XML;
                labelXML.Text = myXCLTrans.XML;
                xmlDoc.Save(@"" + System.Environment.GetEnvironmentVariable("temp") + "\\" + ConfigurationSettings.AppSettings["importNameFile"]);
                                               
                XMLValidator val = new XMLValidator();
                val.Validate(xmlDoc.InnerXml);


                webBrowser1.Navigate(@"" + System.Environment.GetEnvironmentVariable("temp") + "\\" + ConfigurationSettings.AppSettings["importNameFile"]);
                webBrowser1.Visible = true;
                toolStripStatusLabel2.Text = ConfigurationSettings.AppSettings["finishText"];
                button2.Enabled = true;
                fileToolStripMenuItem.DropDownItems[0].Visible = true;
                toolStripProgressBar1.Value = 0;
            }
            catch (Exception ex) {

                MessageBox.Show(ex.Message, ConfigurationSettings.AppSettings["errorText"], MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                toolStripStatusLabel2.Text = ConfigurationSettings.AppSettings["errorText"];
                webBrowser1.DocumentText = ex.Message; 
            }
        }

        private void saveFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button2_Click(sender, e);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //pick whatever filename with .xml extension            
            XCLsaveFileDialog.FileName = XCLFileDialog.FileName + ".XML";
            XCLsaveFileDialog.ShowDialog();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.InnerXml = labelXML.Text;
            xmlDoc.Save(XCLsaveFileDialog.FileName);

            toolStripStatusLabel2.Text = ConfigurationSettings.AppSettings["saveFileText1"] + XCLFileDialog.FileName + ConfigurationSettings.AppSettings["saveFileText2"];
        }

        private void btCreateXMLFlux_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel2.Text = ConfigurationSettings.AppSettings["workingText"]; 

            BdocXCLTransFlux myXCLTrans = new BdocXCLTransFlux(XCLFileDialog.FileName);
            try
            {
                myXCLTrans.createXML(toolStripProgressBar1);
                webBrowser2.DocumentText = myXCLTrans.XML;

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.InnerXml = myXCLTrans.XML;
                labelXML.Text = myXCLTrans.XML;

                xmlDoc.Save(@"" + System.Environment.GetEnvironmentVariable("temp") + "\\" + ConfigurationSettings.AppSettings["fluxNameFile"]);

                //XMLValidator val = new XMLValidator();
                //val.Validate(xmlDoc.InnerXml);

                ;

                webBrowser2.Navigate(@"" + System.Environment.GetEnvironmentVariable("temp") + "\\" + ConfigurationSettings.AppSettings["fluxNameFile"]);

                webBrowser2.Visible = true;
                toolStripStatusLabel2.Text = ConfigurationSettings.AppSettings["finishText"]; ;
                fileToolStripMenuItem.DropDownItems[0].Visible = true;
                toolStripProgressBar1.Value = 0;
                button4.Enabled = true;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, ConfigurationSettings.AppSettings["errorText"], MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                toolStripStatusLabel2.Text = ConfigurationSettings.AppSettings["errorText"];
                webBrowser2.DocumentText = ex.Message;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //pick whatever filename with .xml extension            
            XCLsaveFileDialog.FileName = XCLFileDialog.FileName + ".XML";
            XCLsaveFileDialog.ShowDialog();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.InnerXml = labelXML.Text;
            xmlDoc.Save(XCLsaveFileDialog.FileName);

            toolStripStatusLabel2.Text = ConfigurationSettings.AppSettings["saveFileText1"] + XCLFileDialog.FileName + ConfigurationSettings.AppSettings["saveFileText2"];
        }

        private void aboutBDOCExcel2XMLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 MyAboutBox = new AboutBox1();
            MyAboutBox.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel2.Text = ConfigurationSettings.AppSettings["workingText"];

            BDOCXCLTransTxt myXCLTrans = new BDOCXCLTransTxt (XCLFileDialog.FileName);
            try
            {
                myXCLTrans.createXML(toolStripProgressBar1);

                System.IO.StreamWriter sw = new System.IO.StreamWriter((@"" + System.Environment.GetEnvironmentVariable("temp") + "\\" + ConfigurationSettings.AppSettings["txtNameFile"]), false, Encoding.GetEncoding(1252), 2048);
                //sw.WriteLine("<HTML><HEAD><META http-equiv=Content-Type content=\"text/html; charset=utf-8\"><META content=\"MSHTML 6.00.6000.16705\" name=GENERATOR></HEAD>");

                
                                
                switch (comboBox1.SelectedItem.ToString())
                {
                    case "Data":
                        //webBrowser3.DocumentText = myXCLTrans.FinalTXTData;
                        //Byte[] myBytes = Encoding.ASCII.GetBytes(myXCLTrans.FinalTXTData);
                        //File.WriteAllBytes((@"" + System.Environment.GetEnvironmentVariable("temp") + "\\" + ConfigurationSettings.AppSettings["txtNameFile"]), myBytes);
                        sw.Write(myXCLTrans.FinalTXTData);
                        labelXML.Text = myXCLTrans.FinalTXTData;
                        break;
                    case "Entity":
                        //webBrowser3.DocumentText = myXCLTrans.FinalTXTEntity;
                        sw.Write(myXCLTrans.FinalTXTEntity);
                        labelXML.Text = myXCLTrans.FinalTXTEntity;
                        break;
                    case "Flux":
                        //webBrowser3.DocumentText = myXCLTrans.FinalTXTStream;
                        sw.Write(myXCLTrans.FinalTXTStream);
                        labelXML.Text = myXCLTrans.FinalTXTStream;
                        break;
                    default:
                        //webBrowser3.DocumentText = myXCLTrans.FinalTXTData;
                        sw.Write(myXCLTrans.FinalTXTData);
                        labelXML.Text = myXCLTrans.FinalTXTData;
                        break;
                }
                sw.Flush();
                sw.Close();
                webBrowser3.Navigate(@"" + System.Environment.GetEnvironmentVariable("temp") + "\\" + ConfigurationSettings.AppSettings["txtNameFile"]);

                webBrowser3.Visible = true;
                toolStripStatusLabel2.Text = myXCLTrans.total + ". " + ConfigurationSettings.AppSettings["finishText"];                 
                fileToolStripMenuItem.DropDownItems[0].Visible = true;
                toolStripProgressBar1.Value = 0;
                button5.Enabled = true;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, ConfigurationSettings.AppSettings["errorText"], MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                toolStripStatusLabel2.Text = ConfigurationSettings.AppSettings["errorText"];
                webBrowser3.DocumentText = ex.Message;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            //pick whatever filename with .xml extension            
            XCLsaveFileDialog.FileName = XCLFileDialog.FileName + ".txt";
            XCLsaveFileDialog.ShowDialog();
            System.IO.StreamWriter sw = new System.IO.StreamWriter(XCLsaveFileDialog.FileName, false, Encoding.GetEncoding(1252), 2048);
            sw.Write(labelXML.Text);
            sw.Flush();
            sw.Close();
            
            toolStripStatusLabel2.Text = ConfigurationSettings.AppSettings["saveFileText1"] + XCLFileDialog.FileName + ConfigurationSettings.AppSettings["saveFileText3"];
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (Worksheet mySheet in this.workBook.Sheets)
            {
                //Worksheet mySheet = (Worksheet)workBook.Sheets.get_Item(i);
                comboBox2.Items.Add(mySheet.Name);
            }

            //Worksheet workSheet = (Worksheet)workBook.Sheets..ActiveSheet;
            //comboBox2.SelectedItem = workSheet.Name;
        }

       

               
    
    }
}
