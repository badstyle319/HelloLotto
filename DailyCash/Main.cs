using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Net;
using System.IO;
using System.IO.Compression;
using HtmlAgilityPack;

namespace DailyCash
{
    public partial class Main : Form
    {
        //記錄目前所選資料索引
        static int currentRow = 0;

        //connection連結到資料庫
        //	宣告並設定  連接字串
        const string strConn = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=LT.mdb";
        //	宣告並設定  連接物件conn
        OleDbConnection conn;
        //dataset11物件置於暫時記憶體，以存放查詢結果
        //	宣告並設定 終端機電腦記憶體的暫存物件『datasetNum』
        DataSet datasetNum = new DataSet();
        OleDbDataAdapter adapter;
        int editStatus;//1表示新增資料,2表示修改資料
        //記錄最新一筆資料之日期
        string newestDate = "";
        //舊查詢中之統計數字陣列
        int[] totalNum = new int[49];

        public Main()
        {
            InitializeComponent();

            editStatus = 0;
            messageLabel.Text = "";

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
            ServicePointManager.DefaultConnectionLimit = 50;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                conn = new OleDbConnection(strConn);
                conn.Open();
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
                this.Close();
                return;
            }

            try
            {
                string sql = "create table dailycash (日期 int, " +
                "一 int, 二 int, 三 int, 四 int, 五 int)";
                OleDbCommand cmd = new OleDbCommand(sql, conn);
                cmd.ExecuteNonQuery();
            }
            catch (Exception e1)
            {
                Console.WriteLine(e1.Message);
            }

            RefreshData();
            GotoRow(currentRow);
            UpdateDate();
        }

        private void Main_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (conn.State == ConnectionState.Open)
                conn.Close();
        }

        //更新資料庫內容到異動datagrid上
        private void RefreshData()
        {
            datasetNum.Clear();

            //oledbdataadapter物件建立資料表查詢結果
            //	宣告並設定  查詢『num 資料表』字串
            string str = "select * from dailycash order by 日期";
            //	宣告並設定  資料表查詢物件『adapter』
            adapter = new OleDbDataAdapter(str, conn);

            //	將伺服器資料庫的查詢結果（adapter）存放並填滿到終端機的暫存物件（datasetNum）上的表格"dailycash"
            adapter.Fill(datasetNum, "dailycash");

            //在datagrid來顯示dataset上的資料
            objDV.DataSource = datasetNum.Tables["dailycash"];
            int nRowNum = datasetNum.Tables["dailycash"].Rows.Count;
            if (nRowNum != 0)
                newestDate = datasetNum.Tables["dailycash"].Rows[nRowNum - 1]["日期"].ToString();

            objDV.ClearSelection();
        }

        //跳到第N筆資料方法
        private void GotoRow(int rowNumber)
        {
            objDV.ClearSelection();

            ClearColor(objDV);

            if (datasetNum.Tables["dailycash"].Rows.Count < 1)
                return;

            dateTextBox.Text = datasetNum.Tables["dailycash"].Rows[rowNumber]["日期"].ToString();
            textBox1.Text = datasetNum.Tables["dailycash"].Rows[rowNumber]["一"].ToString();
            textBox2.Text = datasetNum.Tables["dailycash"].Rows[rowNumber]["二"].ToString();
            textBox3.Text = datasetNum.Tables["dailycash"].Rows[rowNumber]["三"].ToString();
            textBox4.Text = datasetNum.Tables["dailycash"].Rows[rowNumber]["四"].ToString();
            textBox5.Text = datasetNum.Tables["dailycash"].Rows[rowNumber]["五"].ToString();
            messageLabel.Text = "第 " + (rowNumber + 1) + " 筆，共有 " + datasetNum.Tables["dailycash"].Rows.Count + " 筆資料。";

            objDV.FirstDisplayedScrollingRowIndex = rowNumber;

            if (rowNumber == 0)
            {
                firstButton.Enabled = false;
                preButton.Enabled = false;
            }
            else
            {
                firstButton.Enabled = true;
                preButton.Enabled = true;
            }

            if (rowNumber >= datasetNum.Tables["dailycash"].Rows.Count - 1)
            {
                nextButton.Enabled = false;
                lastButton.Enabled = false;
            }
            else
            {
                nextButton.Enabled = true;
                lastButton.Enabled = true;
            }
            objDV.Rows[rowNumber].Selected = true;
        }

        //update the newest date for every queries
        private void UpdateDate()
        {
            qDateTextBox2.Text = newestDate;
        }

        //清除cellcolor
        private void ClearColor(DataGridView dgv)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
                for (int j = 0; j < dgv.Columns.Count; j++)
                    dgv.Rows[i].Cells[j].Style.BackColor = Color.Empty;
        }

        //第一筆按鈕
        private void FirstButton_Click(object sender, EventArgs e)
        {
            currentRow = 0;
            GotoRow(currentRow);
        }

        //上一筆按鈕
        private void PreButton_Click(object sender, EventArgs e)
        {
            currentRow -= 1;
            GotoRow(currentRow);
        }

        //下一筆按鈕
        private void NextButton_Click(object sender, EventArgs e)
        {
            currentRow += 1;
            GotoRow(currentRow);
        }

        //最後一筆按鈕
        private void LastButton_Click(object sender, EventArgs e)
        {
            currentRow = objDV.Rows.Count - 1;
            GotoRow(currentRow);
        }

        //啟動編輯模式
        private void EnableEdit()
        {
            dateTextBox.Enabled = true;
            textBox1.Enabled = true;
            textBox2.Enabled = true;
            textBox3.Enabled = true;
            textBox4.Enabled = true;
            textBox5.Enabled = true;
            okButton.Enabled = true;
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            firstButton.Enabled = false;
            preButton.Enabled = false;
            nextButton.Enabled = false;
            lastButton.Enabled = false;
        }

        //關閉編輯模式
        private void DisableEdit()
        {
            dateTextBox.Enabled = false;
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            textBox3.Enabled = false;
            textBox4.Enabled = false;
            textBox5.Enabled = false;
            okButton.Enabled = false;
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            firstButton.Enabled = true;
            preButton.Enabled = true;
            nextButton.Enabled = true;
            lastButton.Enabled = true;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            EnableEdit();
            dateTextBox.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            editStatus = 1;
            dateTextBox.Focus();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            EnableEdit();
            editStatus = 2;
            button1.Enabled = false;
            button3.Enabled = false;
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            DialogResult answer = MessageBox.Show("您確定要刪除此筆資料嗎？", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

            if (answer == DialogResult.Yes)
            {
                //刪除資料庫內的記錄
                //	設定刪除記錄的 SQL語法 及資料庫執行指令OleDbCommand
                string str = "delete * from dailycash where 日期=" + Int32.Parse(dateTextBox.Text) + "";
                OleDbCommand cmd = new OleDbCommand(str, conn);

                //執行資料庫指令OleDbCommand
                cmd.ExecuteNonQuery();

                RefreshData();
                currentRow = 0;
                GotoRow(currentRow);
            }
            UpdateDate();
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            DialogResult answer;

            if (editStatus == 1)
            {
                answer = MessageBox.Show("您確定要新增此筆資料嗎？", "新增資料", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

                if (answer == DialogResult.Yes)
                {
                    //新增記錄到資料庫內
                    //	設定新增記錄的 SQL語法 及資料庫執行指令OleDbCommand
                    string str = "Insert Into dailycash(日期,一,二,三,四,五)Values(" + Int32.Parse(dateTextBox.Text) + "," + Int32.Parse(textBox1.Text) + "," + Int32.Parse(textBox2.Text) + "," + Int32.Parse(textBox3.Text) + "," + Int32.Parse(textBox4.Text) + "," + Int32.Parse(textBox5.Text) + ")";

                    OleDbCommand cmd = new OleDbCommand(str, conn);

                    //執行資料庫指令OleDbCommand
                    cmd.ExecuteNonQuery();

                    RefreshData();
                    currentRow = 0;
                    GotoRow(currentRow);
                }
            }
            else if (editStatus == 2)
            {
                answer = MessageBox.Show("您確定要修改此筆資料嗎？", "修改資料", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

                if (answer == DialogResult.Yes)
                {
                    //修改資料庫內的記錄
                    //	設定修改記錄的  SQL語法及資料庫執行指令OleDbCommand
                    string str = "Update dailycash set 日期 = " + Int32.Parse(dateTextBox.Text) + ",一=" + Int32.Parse(textBox1.Text) + ",二=" + Int32.Parse(textBox2.Text) + ",三=" + Int32.Parse(textBox3.Text) + ",四=" + Int32.Parse(textBox4.Text) + ",五=" + Int32.Parse(textBox5.Text) + " where 日期= " + Int32.Parse(dateTextBox.Text) + "";

                    OleDbCommand cmd = new OleDbCommand(str, conn);

                    //執行資料庫指令OleDbCommand
                    cmd.ExecuteNonQuery();

                    RefreshData();
                    GotoRow(currentRow);
                }
            }
            DisableEdit();
            UpdateDate();
        }

        private void BtnCrawl_Click(object sender, EventArgs e)
        {
            string urlAddress = "https://www.taiwanlottery.com.tw/Lotto/Dailycash/history.aspx";

            var request = (HttpWebRequest)WebRequest.Create(urlAddress);
            var str = "";

            request.ServicePoint.Expect100Continue = false;
            request.ServicePoint.UseNagleAlgorithm = false;
            request.AllowAutoRedirect = false;
            request.AllowWriteStreamBuffering = false;
            request.Headers.Add(HttpRequestHeader.AcceptEncoding, "gzip,deflate");
            request.Accept = "*/*";
            request.ContentType = "application/x-www-form-urlencoded";
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.106 Safari/537.36";
            request.Timeout = 5000;
            request.Method = "GET";

            using (var response = (HttpWebResponse)request.GetResponse())
            {
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    var strContentEncoding = response.ContentEncoding.ToLower();
                    if (strContentEncoding.Contains("gzip"))
                    {
                        Console.WriteLine("gzip stream:");
                        using (GZipStream stream = new GZipStream(response.GetResponseStream(), CompressionMode.Decompress))
                        {
                            using (StreamReader sr = new StreamReader(stream, System.Text.Encoding.UTF8))
                            {
                                str = sr.ReadToEnd();
                            }
                        }
                    }
                    else if (strContentEncoding.Contains("deflate"))
                    {
                        Console.WriteLine("deflate stream:");
                        using (DeflateStream stream = new DeflateStream(response.GetResponseStream(), CompressionMode.Decompress))
                        {
                            using (StreamReader sr = new StreamReader(stream, System.Text.Encoding.UTF8))
                            {
                                str = sr.ReadToEnd();
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("normal stream:");
                        using (Stream stream = response.GetResponseStream())
                        {
                            using (StreamReader sr = new StreamReader(stream, System.Text.Encoding.UTF8))
                            {
                                str = sr.ReadToEnd();
                            }
                        }
                    }
                    Console.WriteLine(str);

                    //Stream receiveStream = response.GetResponseStream();
                    //StreamReader readStream = null;

                    //if (String.IsNullOrEmpty(response.CharacterSet))
                    //    readStream = new StreamReader(receiveStream);
                    //else
                    //    readStream = new StreamReader(receiveStream, Encoding.GetEncoding(response.CharacterSet));

                    //string data = readStream.ReadToEnd();
                    //Console.WriteLine(data);
                    //response.Close();
                    //readStream.Close();
                }
            }
        }
    }
}
