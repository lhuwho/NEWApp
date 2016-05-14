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
using System.Security.Permissions;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
       
      
        public Form1()
        {
            InitializeComponent();
            button3.Hide();
            progressBar1.Hide();
        }

        List<ExcelRowData> ExcelRowDataList=new List<ExcelRowData>();
        Microsoft.Office.Interop.Excel.Application xlApp = null;
        Workbook wb = null;
        Worksheet ws = null;
        Range aRange = null;
        List<string> ColorList = new List<string>();
        Connect_Net z = new Connect_Net();
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
                        getColors();
                        SaveOrInsertSheet(openFileDialog1.FileName, (Worksheet)xlApp.Worksheets[i]);
                    }
                    label4.Text = "共    " + ExcelRowDataList.Count() + "   個檔案待傳輸";
                    button3.Show();
                    
                    
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


        public class Connect_Net
        {
            public string Server_IP
            {
                get;
                set;
            }
            public string UserNmae
            {
                get;
                set;
            }
            public string Password
            {
                get;
                set;
            }
        }

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool LogonUser(
            string lpszUsername,
            string lpszDomain,
            string lpszPassword,
            int dwLogonType,
            int dwLogonProvider,
            ref IntPtr phToken);

        //登出
        [DllImport("kernel32.dll")]
        public extern static bool CloseHandle(IntPtr hToken);


        public List<string> getFTPList()
        {
            List<string> strList = new List<string>();

            if (textBox1.Text.Length > 0) { 
         
                z.UserNmae = textBox2.Text;
                z.Password = textBox3.Text;
                z.Server_IP = textBox1.Text;


                string newftpUrl = z.Server_IP;// +"tote_checkline";
                string IPath = String.Format(@"\\{0}", newftpUrl);
                const int LOGON32_PROVIDER_DEFAULT = 0;
                const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
                IntPtr tokenHandle = new IntPtr(0);
                tokenHandle = IntPtr.Zero;
                try
                {
                    bool returnValue = LogonUser(z.UserNmae, z.Server_IP, z.Password,
                    LOGON32_LOGON_NEW_CREDENTIALS,
                    LOGON32_PROVIDER_DEFAULT,
                    ref tokenHandle);
                    WindowsIdentity w = new WindowsIdentity(tokenHandle);
                    System.Security.Principal.WindowsImpersonationContext kk = w.Impersonate();
                    if (false == returnValue)
                    {
                        strList.Add("無法連線");
                    }

                    DirectoryInfo dir = new DirectoryInfo(IPath );
                    if (dir.Exists != true)
                    {
                        dir = new DirectoryInfo(IPath);
                    }
                    //FileInfo[] inf = dir.GetFiles();
                    FileSystemInfo[] inf = dir.GetFileSystemInfos();
                    
                    for (int i = 0; i < inf.Length; i++)
                    {
                        strList.Add(inf[i].Name);
                        //Console.WriteLine(inf[i].Name);
                        //檔案取回
                        // System.IO.File.Copy("\\\\" + z.Server_IP + "\\ftp\\" + inf[i].Name, "C:\\" + inf[i].Name + ".jpg");
                    }
                    kk.Undo();
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
            //List<string> strList = new List<string>();
            //if (textBox1.Text.Length > 0)
            //{

            //    FtpWebRequest f = (FtpWebRequest)WebRequest.Create(new Uri("ftp://" + textBox1.Text));
            //    f.Method = WebRequestMethods.Ftp.ListDirectory;
            //    f.UseBinary = true;
            //    f.AuthenticationLevel = System.Net.Security.AuthenticationLevel.MutualAuthRequested;
            //    f.Credentials = new NetworkCredential(textBox2.Text, textBox3.Text);

            //    try
            //    {
            //        StreamReader sr = new StreamReader(f.GetResponse().GetResponseStream());
            //        string str = sr.ReadLine();
            //        while (str != null)
            //        {
            //            strList.Add(str);
            //            str = sr.ReadLine();

            //        }

            //        sr.Close();
            //        sr.Dispose();
            //        f = null;
            //    }
            //    catch (Exception e)
            //    {
            //        strList.Add(e.Message.ToString());
            //    }
            //}
            //else
            //{
            //    strList.Add("連線失敗");
            //}

            //return strList;
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


                if (ColorList.Contains(theRow.Color.ToLower()))
                {
                    ExcelRowDataList.Add(theRow);
                }
                




                //往下抓一列Excel範圍
                rowIndex++;
                this.aRange = ws.get_Range("A" + rowIndex.ToString(), "G" + rowIndex.ToString());
            }


        }
        #endregion
        private void getColors()
        {
            ColorList.Clear();
            if (checkedListBox1.CheckedItems.Count != 0)
            {
           
                for (int x = 0; x <= checkedListBox1.CheckedItems.Count - 1; x++)
                {
               
                    ColorList.Add(checkedListBox1.CheckedItems[x].ToString().ToLower());
                }
             
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            this.backgroundWorker1.WorkerReportsProgress = true;
           
            for (int i = 0; i < ExcelRowDataList.Count; i++)
            {
           
                File_DownLoad(ExcelRowDataList[i]);
                //int barValue = (100 / ExcelRowDataList.Count) * (i);
                backgroundWorker1.ReportProgress((i+1));
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            label4.Text = string.Format("共有  {1}  個檔案，正在下載第  {0}  個 。", e.ProgressPercentage, ExcelRowDataList.Count);
            progressBar1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            label4.Text = string.Format("下載完成。 ");
            progressBar1.Value = ExcelRowDataList.Count;
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            progressBar1.Show();
            progressBar1.Maximum = ExcelRowDataList.Count;
            label4.Text = "下載檔案中。";
            backgroundWorker1.RunWorkerAsync();
        }

        private void File_DownLoad(ExcelRowData rowData)
        {
            FTPgetFile();
            string newftpUrl =z.Server_IP + rowData.Item;
            string IPath = String.Format(@"\\{0}", newftpUrl);
            string extension = ".png";

            try
            {

                DirectoryInfo dir = new DirectoryInfo(IPath + "\\");
                while (dir.Exists != true)
                {
                    dir = new DirectoryInfo(IPath);
                }

                FileInfo[] inf = dir.GetFiles(rowData.FileName + extension);
                if (!(dir.Exists && inf.Length > 0 && inf[0].Exists))
                {

                    newftpUrl = textBox1.Text + rowData.Item;
                    IPath = String.Format(@"\\{0}", newftpUrl);
                    DirectoryInfo dir_checkagain = new DirectoryInfo(IPath + "\\");
                    inf = dir_checkagain.GetFiles(rowData.FileName + extension);
                }
                if (dir.Exists && inf.Length > 0 && inf[0].Exists)
                {
                    string pacth =  System.Windows.Forms.Application.StartupPath+"\\";
                    if (!System.IO.Directory.Exists(pacth)) //檢查文件夾是否存在。
                    {
                        System.IO.Directory.CreateDirectory(pacth); //不存在，創建資料夾。
                    }

                    string oldName = inf[0].Name;
                    string ReName = oldName;
                    string ReNameItem = rowData.FileName + "_item" + extension;
                    if (oldName.IndexOf("_") > 0)
                    {
                        ReName = "images";
                        ReNameItem = oldName.Substring(0, oldName.IndexOf("_")) + "_item" + extension;
                    }               

               
                    Bitmap currentPicture = new Bitmap(inf[0].DirectoryName + "\\" + inf[0].Name);
                    this.ResizePicFroMM(currentPicture, pacth + ReName, 300);

                }
                else
                {
               

                }

            }
            catch (Exception e)
            {
             
            }
        }

        private string[] ResizePicFroMM(System.Drawing.Bitmap pinsrc, string sfileName, double targetMM_height)
        {
            string[] returnValue = new string[2];
            returnValue[0] = "0";
            returnValue[1] = "";
            try
            {
                int Width = pinsrc.Width;
                int Height = pinsrc.Height;
                int new_height, new_width;

                //double target_height = (targetMM_height / 25.4) * pinsrc.HorizontalResolution;

                /*if (Height < target_height)
                {
                    new_width = (int)(target_height / Height) * Width;
                    new_height = (int)target_height;
                }
                else
                {*/
                new_width = Width;
                new_height = Height;
                //}

                System.Drawing.Bitmap resizeIMG = Resize(pinsrc, new_width, new_height);
                double XYDpi = Math.Round(pinsrc.HorizontalResolution, 1);
                resizeIMG.SetResolution((float)XYDpi, (float)XYDpi);
                resizeIMG.Save(sfileName, System.Drawing.Imaging.ImageFormat.Png);
                resizeIMG.Dispose();
                pinsrc.Dispose();
                returnValue[0] = "1";
            }
            catch (Exception e)
            {
                returnValue[0] = "-1";
                returnValue[1] = e.Message;
            }
            finally
            {
                pinsrc.Dispose();
            }
            return returnValue;
        }

        private System.Drawing.Bitmap Resize(System.Drawing.Bitmap src, int resizewidth, int resizeheight)
        {
            System.Drawing.Bitmap resizeb = new System.Drawing.Bitmap(resizewidth, resizeheight);
            System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(resizeb);
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            g.Clear(System.Drawing.Color.Transparent);
            g.DrawImage(src, new System.Drawing.Rectangle(0, 0, resizewidth, resizeheight), new System.Drawing.Rectangle(0, 0, src.Width, src.Height), System.Drawing.GraphicsUnit.Pixel);

            g.Dispose();

            return resizeb;
        }

        public void  FTPgetFile()
	{
		//
		// TODO: Add constructor logic here
		//
        
        const int LOGON32_PROVIDER_DEFAULT = 0;
        const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
        IntPtr tokenHandle = new IntPtr(0);
        tokenHandle = IntPtr.Zero;
        try
        {
            bool LogonExit = LogonUser(z.UserNmae, z.Server_IP, z.Password,
            LOGON32_LOGON_NEW_CREDENTIALS,
            LOGON32_PROVIDER_DEFAULT,
            ref tokenHandle);

            WindowsIdentity w = new WindowsIdentity(tokenHandle);
            w.Impersonate();
            if (false == LogonExit)
            {
         
            }
        }
        catch (Exception e)
        {
        
        }
	}

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {          
            try
            {
                backgroundWorker1.CancelAsync();
                backgroundWorker1.Dispose();
                //刪除 Windows工作管理員中的Excel.exe 處理緒.
                if (this.xlApp != null)
                {
                    xlApp.Workbooks.Close();
                    xlApp.Quit();
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
