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

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }



        private void button2_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("1");
            //FtpWebRequest ftpReq;
            richTextBox1.Text = "";
            List<string> ListResult = new List<string>();
            ListResult = getFTPList();

            int count = 0;
            foreach (var LR in ListResult) {
                count++;
                richTextBox1.Text += ("項目 " + count + ":  " + LR + "\r\n");
            }

            //ListResult = getFTPList();
          
        }

        public List<string> getFTPList()
        {
            List<string> strList = new List<string>();
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
            catch (Exception e) {
                strList.Add(e.Message.ToString());
            }

            return strList;
        }
    }
}
