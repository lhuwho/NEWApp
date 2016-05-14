using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
       
      
        public Form1()
        {
            InitializeComponent();
        }

        List<ExcelRowData> ExcelRowDataList=new List<ExcelRowData>();
        Microsoft.Office.Interop.Excel.Application xlApp = null;
        Workbook wb = null;
        Worksheet ws = null;
        Range aRange = null;
        private void button1_Click(object sender, EventArgs e)
        {
            //openFileDialog1
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {

                   
                    if (this.xlApp == null)
                    {
                        this.xlApp = new Microsoft.Office.Interop.Excel.Application();
                    }
                    this.xlApp.Workbooks.Open(openFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    this.wb = xlApp.Workbooks[1];//第一個Workbook
                    this.wb.Save();
                    for (int i = 1; i <= xlApp.Worksheets.Count; i++)
                    {
                        SaveOrInsertSheet(openFileDialog1.FileName, (Worksheet)xlApp.Worksheets[i]);
                    }
                    label4.Text = "共    " + ExcelRowDataList.Count() + "   個檔案待傳輸";
                      
                 
                    
                }
                catch 
                {
                   
                }
                finally
                {
                    xlApp.Workbooks.Close();
                    xlApp.Quit();
                    try
                    {
                        //刪除 Windows工作管理員中的Excel.exe 處理緒.
                        if (this.xlApp != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(this.xlApp);
                        }
                      
                        if (this.ws != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(this.ws);
                        }
                        if (this.aRange != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(this.aRange);                        
                        }
                      
                    }
                    catch { }
                    this.xlApp = null;
                    this.wb = null;
                    this.ws = null;
                    this.aRange = null;
                    GC.Collect();
                }
            }
        }



        private void button2_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("1");
            //FtpWebRequest ftpReq;
            richTextBox1.Text = "";
            List<string> ListResult = new List<string>();
            ListResult = getFTPList();

            int count = 0;
            foreach (var LR in ListResult)
            {
                count++;
                richTextBox1.Text += ("項目 " + count + ":  " + LR + "\r\n");
            }

            //ListResult = getFTPList();

        }

        public List<string> getFTPList()
        {
            List<string> strList = new List<string>();
            if (textBox1.Text.Length > 0)
            {

                FtpWebRequest f = (FtpWebRequest)WebRequest.Create(new Uri("ftp://" + textBox1.Text));
                f.Method = WebRequestMethods.Ftp.ListDirectory;
                f.UseBinary = true;
                f.AuthenticationLevel = System.Net.Security.AuthenticationLevel.MutualAuthRequested;
                f.Credentials = new NetworkCredential(textBox2.Text, textBox3.Text);

                try
                {
                    StreamReader sr = new StreamReader(f.GetResponse().GetResponseStream());
                    string str = sr.ReadLine();
                    while (str != null)
                    {
                        strList.Add(str);
                        str = sr.ReadLine();

                    }

                    sr.Close();
                    sr.Dispose();
                    f = null;
                }
                catch (Exception e)
                {
                    strList.Add(e.Message.ToString());
                }
            }
            else
            {
                strList.Add("連線失敗");
            }

            return strList;
        }

     

        #region 把Excel資料Insert into Table
        private void SaveOrInsertSheet(string excel_filename, Worksheet ws)
        {

            //要開始讀取的起始列(微軟Worksheet是從1開始算)
            int rowIndex = 2;
            //取得一列的範圍
            this.aRange = ws.get_Range("A" + rowIndex.ToString(), "G" + rowIndex.ToString());
            //判斷Row範圍裡第1格有值的話，迴圈就往下跑
            while (((object[,])this.aRange.Value2)[1, 1] != null)//用this.aRange.Cells[1, 1]來取值的方式似乎會造成無窮迴圈？
            {

                ExcelRowData theRow = new ExcelRowData();
                theRow.Sheep = ws.Name;
                //範圍裡第1格的值
                theRow.Item = ((object[,])this.aRange.Value2)[1, 1] != null ? ((object[,])this.aRange.Value2)[1, 1].ToString() : "";

                //範圍裡第2格的值
                theRow.FileName = ((object[,])this.aRange.Value2)[1, 2] != null ? ((object[,])this.aRange.Value2)[1, 2].ToString() : "";

                //範圍裡第4格的值
                theRow.Color = ((object[,])this.aRange.Value2)[1, 4] != null ? ((object[,])this.aRange.Value2)[1, 4].ToString() : "";

                theRow.Size = ((object[,])this.aRange.Value2)[1, 5] != null ? ((object[,])this.aRange.Value2)[1, 5].ToString() : "";

                theRow.Quantity = ((object[,])this.aRange.Value2)[1, 7] != null ? ((object[,])this.aRange.Value2)[1, 7].ToString() : "";



                ExcelRowDataList.Add(theRow);




                //往下抓一列Excel範圍
                rowIndex++;
                this.aRange = ws.get_Range("A" + rowIndex.ToString(), "G" + rowIndex.ToString());
            }


        }
        #endregion
    }
}
