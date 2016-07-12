using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelTest2
{
    public partial class Form1 : Form
    {
        String pathFile = "";
        String pathDirection = "";
        String pathOutputDirection = "";
        String fileName = "";

        String[] headerTable = { "Status", "ODM", "Vendor","RMA No.","AcerP/N",
            "QTY","ODM U/P","Vendor C/N","Vendor U/P","MM#/ID"};

        StringBuilder sbResult = new StringBuilder();

        Excel.Application app = null;
        Excel.Workbook wbks;

        public Form1()
        {
            InitializeComponent();
            this.Text += " Ver0.1.160708.2";
        }
        void initailExcel()
        {
            //檢查PC有無Excel在執行
            bool flag = false;
            foreach (var item in Process.GetProcesses())
            {
                if (item.ProcessName == "EXCEL")
                {
                    flag = true;
                    break;
                }
            }

            if (!flag)
            {
                this.app = new Excel.Application();
            }
            else
            {
                object obj = Marshal.GetActiveObject("Excel.Application");//引用已在執行的Excel
                app = obj as Excel.Application;
            }

            this.app.Visible = false;//設false效能會比較好
        }
        public void fileProcess()
        {
            sbResult.Clear();

            object oMissiong = System.Reflection.Missing.Value;
            //Creat new folder
            pathOutputDirection = pathDirection + @"\Result_" + fileName;
            if (!Directory.Exists(pathOutputDirection))
            {
                Directory.CreateDirectory(pathOutputDirection);
            }

            try
            {
                //Open 
                initailExcel();
                //Open File
                wbks = app.Workbooks.Open(pathFile, oMissiong,
                    XlFileAccess.xlReadOnly, oMissiong, oMissiong, oMissiong,
                    oMissiong, oMissiong, oMissiong, oMissiong, oMissiong,
                    oMissiong, oMissiong, oMissiong, oMissiong);
                Excel.Sheets sheets = wbks.Sheets;
                Thread tt;
                foreach (Excel.Worksheet sheet in sheets)
                {
                    //if (sheet.Name.Equals("Battery"))
                    //{
                    //    sheetProcess(sheet, "Battery");
                    //    //saveSheetToCSV(sheet);
                    //}
                    //if (!sheet.Name.Equals("DT HDD"))
                    //{
                    //    sheetProcess(sheet, sheet.Name);
                    //}

                    sheetProcess(sheet, sheet.Name);
                }
            }
            finally
            {
                wbks.Saved = true;
                wbks.Close();
                wbks = null;
                //app.Quit();
                //app = null;
                GC.Collect();
                if (sbResult.ToString().Equals(""))
                {
                    MessageBox.Show("所有Sheet处理完毕~","没问题~");
                }
                else
                {
                    MessageBox.Show(sbResult.ToString(),"有问题……");
                }
            }
            //app.Quit();
            //app = null;
        }

        //public void sheetProcessTask(object sheet)
        //{
        //    Excel.Worksheet Sheet = sheet as Excel.Worksheet;
        //    sheetProcess(Sheet, Sheet.Name);
        //}

        public void sheetProcess(Excel.Worksheet Sheet, String sheetName)
        {
            Excel.Workbook oWB = (Excel.Workbook)(app.Workbooks.Add(Missing.Value));
            Excel.Worksheet oSheet = (Excel.Worksheet)oWB.ActiveSheet;
            try
            {
                Excel.Range range = Sheet.get_Range("A1", "Z1");
                String strTemp = "";
                int rowMax = 5000;
                int[] rowIndex = new int[10];
                //find Index
                foreach (Excel.Range item in range)
                {
                    strTemp = item.Cells.Text;
                    if (strTemp.Equals("Status"))
                    {
                        strTemp = Sheet.Cells[2, item.Column].Text;
                        if (!strTemp.Equals(""))
                        {
                            //set Index
                            rowIndex[0] = item.Column;
                            //find RowMax
                            for (int i = 1; i < rowMax; i++)
                            {
                                strTemp = Sheet.Cells[i, item.Column].Text;
                                if (strTemp.Equals(""))
                                {
                                    rowMax = i;
                                }
                            }
                        }
                    }
                    if (strTemp.Equals("ODM"))
                    {
                        //set Index
                        rowIndex[1] = item.Column;
                    }
                    if (strTemp.Equals("Vendor"))
                    {
                        //set Index
                        rowIndex[2] = item.Column;
                    }
                    if (strTemp.Equals("RMA No."))
                    {
                        //set Index
                        rowIndex[3] = item.Column;
                    }
                    if (strTemp.Equals("AcerP/N"))
                    {
                        //set Index
                        rowIndex[4] = item.Column;
                    }
                    if (strTemp.Equals("QTY"))
                    {
                        //set Index
                        rowIndex[5] = item.Column;
                    }
                    if (strTemp.Equals("ODM U/P"))
                    {
                        //set Index
                        rowIndex[6] = item.Column;
                    }
                    if (strTemp.Equals("Vendor C/N"))
                    {
                        //set Index
                        rowIndex[7] = item.Column;
                    }
                    if (strTemp.Equals("Vendor U/P"))
                    {
                        //set Index
                        rowIndex[8] = item.Column;
                    }
                    if (strTemp.Equals("MM#/ID"))
                    {
                        //set Index
                        rowIndex[9] = item.Column;
                    }
                }

                int result = 0;
                for (int t = 0; t < 9; t++)
                {
                    if (rowIndex[t] == 0)
                    {
                        result = 1; 
                    }
                }
                //Error Message
                if (result == 1)
                {
                    sbResult.Append(sheetName + " Error: ");
                    for (int t = 0; t < 9; t++)
                    {
                        if (rowIndex[t] == 0)
                        {
                            result = 1;
                            String str = " No " + headerTable[t] + ",";
                            sbResult.Append(str);
                        }
                    }
                    sbResult.Append("\r\n\r\n");
                }
                
                if (result == 0)
                {
                    //copy to excel
                    int j = 2;
                    //Header
                    for (int k = 1; k < 11; k++)
                    {
                        oSheet.Cells[1, k] = Sheet.Cells[1, rowIndex[k - 1]];
                    }
                    //fill Data
                    for (int i = 2; i < rowMax; i++)
                    {
                        Excel.Range orange = Sheet.Cells[i, rowIndex[0]];
                        int colorIndex = (int)orange.Interior.Color;
                        Color testColor = Color.FromArgb(colorIndex);
                        strTemp = orange.Text;
                        strTemp = strTemp.ToLower();

                        if (strTemp.Contains("open") && testColor.Name.Equals("ffff"))
                        {
                            strTemp = Sheet.Cells[i, rowIndex[7]].Text;
                            if (!strTemp.Equals(""))
                            {
                                for (int k = 1; k < 11; k++)
                                {
                                    oSheet.Cells[j, k] = Sheet.Cells[i, rowIndex[k - 1]];
                                }
                                j++;
                            }
                        }
                    }
                    oSheet.get_Range("A1", "Z500").EntireColumn.AutoFit();
                    oSheet.Name = sheetName;
                    //new Sheet
                    Excel.Worksheet r1Sheet = (Excel.Worksheet)oWB.Sheets.Add();
                    r1Sheet.Name = sheetName + "_ODM_CSV";
                    Excel.Worksheet r2Sheet = (Excel.Worksheet)oWB.Sheets.Add();
                    r2Sheet.Name = sheetName + "_Vendor_CSV";
                    //parase Data
                    //Sort Vendor C\N
                    oSheet.get_Range("A2", "Z500").Sort(oSheet.Cells[1, 8], XlSortOrder.xlAscending,
                        Missing.Value, Missing.Value, XlSortOrder.xlAscending, Missing.Value, XlSortOrder.xlAscending,
                        XlYesNoGuess.xlNo, Missing.Value, XlSortOrientation.xlSortColumns);
                    //Sort Vendor
                    String strVendorCN = oSheet.Cells[2, 8].Text;
                    String strVendor = oSheet.Cells[2, 3].Text;
                    int preRow1 = 2;
                    int preRow2 = 2;
                    String strStart = "";
                    String strEnd = "";
                    for (int i = 2; i < j + 1; i++)
                    {
                        if (!oSheet.Cells[i, 8].Text.Equals(strVendorCN))
                        {
                            //Sort Vendor
                            strStart = "A" + preRow1.ToString();
                            if (i > 2)
                            {
                                strEnd = "Z" + (i - 1).ToString();

                                oSheet.get_Range(strStart, strEnd).Sort(oSheet.Cells[preRow1, 3], XlSortOrder.xlAscending,
                                Missing.Value, Missing.Value, XlSortOrder.xlAscending, Missing.Value, XlSortOrder.xlAscending,
                                XlYesNoGuess.xlNo, Missing.Value, XlSortOrientation.xlSortColumns);

                                preRow2 = preRow1;

                                for (int d = preRow1; d < i; d++)
                                {
                                    if (!oSheet.Cells[i, 3].Text.Equals(strVendor))
                                    {
                                        //Sort ODM
                                        strStart = "A" + preRow2.ToString();
                                        if (d > 2)
                                        {
                                            strEnd = "Z" + (d - 1).ToString();
                                            oSheet.get_Range(strStart, strEnd).Sort(oSheet.Cells[preRow2, 2], XlSortOrder.xlAscending,
                                            Missing.Value, Missing.Value, XlSortOrder.xlAscending, Missing.Value, XlSortOrder.xlAscending,
                                            XlYesNoGuess.xlNo, Missing.Value, XlSortOrientation.xlSortColumns);
                                            preRow2 = d;
                                        }
                                    }
                                }
                            }
                            preRow1 = i;
                        }
                    }
                    //if the sheet is not empty
                    if (!oSheet.Cells[2, 1].Text.Equals(""))
                    {
                        //make CSV data
                        makeCSVData(oSheet, j - 1, r1Sheet, r2Sheet);

                        //save for Check
                        if (chbSaveProcessed.CheckState == CheckState.Checked)
                        {
                            object oMissiong = System.Reflection.Missing.Value;
                            String pathSheet = pathOutputDirection + @"\" + Sheet.Name + "_Processed" + @".xlsx";
                            String tt = oWB.FileFormat.ToString();
                            oWB.SaveAs(pathSheet, XlFileFormat.xlOpenXMLWorkbook,
                                oMissiong, oMissiong, oMissiong, oMissiong,
                                XlSaveAsAccessMode.xlExclusive, oMissiong, oMissiong, oMissiong);
                        }
                        //save CSV
                        saveSheetToCSV(r1Sheet, r2Sheet);
                    }
                }
            }
            finally
            {
                oWB.Saved = true;
                oWB.Close();
                oWB = null;
            }
        }

        public void makeCSVData(Excel.Worksheet sSheet, int rowMax, Excel.Worksheet t1Sheet, Excel.Worksheet t2Sheet)
        {
            String strODM = "";
            String strVendor = "";
            String strCN = "";
            int j = 1;
            int index = 2;

            #region Sheet1
            //Sheet1
            j = 1;
            index = 2;
            strODM = sSheet.Cells[2, 2].Text;
            strVendor = sSheet.Cells[2, 3].Text;
            strCN = sSheet.Cells[2, 8].Text;
            for (int i = 2; i < rowMax + 2; i++)
            {
                //add to~from message
                if (!strODM.Equals(sSheet.Cells[i, 2].Text)
                    || !strVendor.Equals(sSheet.Cells[i, 3].Text)
                    || !strCN.Equals(sSheet.Cells[i, 8].Text)
                    || i == rowMax + 1)
                {
                    t1Sheet.Cells[j++, 1] = strVendor + " to " + strODM;
                    //Add data
                    for (int k = index; k < i; k++)
                    {
                        t1Sheet.Cells[j, 1] = sSheet.Cells[k, 4].Text;
                        t1Sheet.Cells[j, 2] = sSheet.Cells[k, 5].Text;
                        t1Sheet.Cells[j, 3] = sSheet.Cells[k, 6].Text;
                        //numQty = int.Parse(sSheet.Cells[k, 6].Text);
                        t1Sheet.Cells[j, 4] = "*";
                        t1Sheet.Cells[j, 5] = sSheet.Cells[k, 7].Text;
                        t1Sheet.Cells[j, 5].NumberFormat = "$0.00";
                        t1Sheet.Cells[j, 6] = "=";
                        t1Sheet.Cells[j, 7].Formula = "=C" + j.ToString() + "*" + "E" + j.ToString();
                        t1Sheet.Cells[j, 7].NumberFormat = "$0.00";
                        //add MMID
                        t1Sheet.Cells[j, 8] = sSheet.Cells[k, 10].Text;
                        j++;
                    }
                    t1Sheet.Cells[j, 5] = "TOTAL";
                    t1Sheet.Cells[j, 7].Formula = "=SUM(G" + (j - (i - index) - 1).ToString() + ":G" + (j - 1).ToString() + ")";
                    t1Sheet.Cells[j, 7].NumberFormat = "$0.00";
                    index = i;
                    j += 2;

                    strODM = sSheet.Cells[i, 2].Text;
                    strVendor = sSheet.Cells[i, 3].Text;
                    strCN = sSheet.Cells[i, 8].Text;
                }
            }
            //Align to Right
            t1Sheet.get_Range("A1", "Z100").EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            //Aute fit
            t1Sheet.get_Range("A1", "Z100").EntireColumn.AutoFit();
            #endregion

            #region Sheet2
            //Sheet2
            j = 1;
            index = 2;
            strODM = sSheet.Cells[2, 2].Text;
            strVendor = sSheet.Cells[2, 3].Text;
            strCN = sSheet.Cells[2, 8].Text;
            for (int i = 2; i < rowMax + 2; i++)
            {
                //add to~from message
                if (!strODM.Equals(sSheet.Cells[i, 2].Text)
                    || !strVendor.Equals(sSheet.Cells[i, 3].Text)
                    || !strCN.Equals(sSheet.Cells[i, 8].Text)
                    || i == rowMax + 1)
                {
                    t2Sheet.Cells[j++, 1] = strVendor + " to " + strODM;
                    //Add data
                    for (int k = index; k < i; k++)
                    {
                        t2Sheet.Cells[j, 1] = sSheet.Cells[k, 4].Text;
                        t2Sheet.Cells[j, 2] = sSheet.Cells[k, 5].Text;
                        t2Sheet.Cells[j, 3] = sSheet.Cells[k, 6].Text;
                        //numQty = int.Parse(sSheet.Cells[k, 6].Text);
                        t2Sheet.Cells[j, 4] = "*";
                        t2Sheet.Cells[j, 5] = sSheet.Cells[k, 9].Text;
                        t2Sheet.Cells[j, 5].NumberFormat = "$0.00";
                        t2Sheet.Cells[j, 6] = "=";
                        t2Sheet.Cells[j, 7].Formula = "=C" + j.ToString() + "*" + "E" + j.ToString();
                        t2Sheet.Cells[j, 7].NumberFormat = "$0.00";
                        //add MMID
                        t2Sheet.Cells[j, 8] = sSheet.Cells[k, 10].Text;
                        j++;
                    }
                    t2Sheet.Cells[j, 5] = "TOTAL";
                    t2Sheet.Cells[j, 7].Formula = "=SUM(G" + (j - (i - index) - 1).ToString() + ":G" + (j - 1).ToString() + ")";
                    t2Sheet.Cells[j, 7].NumberFormat = "$0.00";
                    index = i;
                    j += 2;

                    strODM = sSheet.Cells[i, 2].Text;
                    strVendor = sSheet.Cells[i, 3].Text;
                    strCN = sSheet.Cells[i, 8].Text;
                }
            }
            //Align to Right
            t2Sheet.get_Range("A1", "Z100").EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            //Aute fit
            t2Sheet.get_Range("A1", "Z100").EntireColumn.AutoFit();
            #endregion
        }

        public void saveSheetToCSV(Excel.Worksheet Sheet1, Excel.Worksheet Sheet2)
        {
            object oMissiong = System.Reflection.Missing.Value;
            //make a new csv path use the new dir and sheet name
            String pathSheet = "";
            pathSheet = pathOutputDirection + @"\" + Sheet1.Name + @".csv";
            Sheet1.SaveAs(pathSheet, XlFileFormat.xlCSV,
                oMissiong, oMissiong, oMissiong, oMissiong,
                oMissiong, oMissiong, oMissiong, oMissiong);
            pathSheet = pathOutputDirection + @"\" + Sheet2.Name + @".csv";
            Sheet2.SaveAs(pathSheet, XlFileFormat.xlCSV,
                oMissiong, oMissiong, oMissiong, oMissiong,
                oMissiong, oMissiong, oMissiong, oMissiong);
        }
        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileOpen = new OpenFileDialog();
            fileOpen.Filter = "Excel|*.xlsx";
            if (fileOpen.ShowDialog() == DialogResult.OK)
            {
                if (fileOpen.CheckFileExists)
                {
                    //get file path
                    pathFile = fileOpen.FileName;
                    //get Direction
                    pathDirection = System.IO.Path.GetDirectoryName(pathFile);
                    //get file name
                    fileName = System.IO.Path.GetFileNameWithoutExtension(pathFile);
                    //debug output file name
                    System.Diagnostics.Debug.WriteLine("\r\nDebug Output:\r\n" + pathFile + "\r\n");
                    //process the file
                    fileProcess();
                }
                else
                {
                    MessageBox.Show("没有文件");
                }
            }
        }
    }
}
