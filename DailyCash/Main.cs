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
#if (NET45 || NET48)
using HtmlAgilityPack;
#endif
using System.Collections.Specialized;

namespace DailyCash
{
    public partial class Main : Form
    {
        const string DB_CONN = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=LT.mdb";

        //記錄最新一筆資料之日期
        string newestDate = "";

        //記錄目前所選資料索引
        static int currentRow = 0;

        int editStatus;//1表示新增資料,2表示修改資料

        //舊查詢中之統計數字陣列
        int[] totalNum = new int[49];

        OleDbConnection conn;
        //dataset物件置於暫時記憶體，以存放查詢結果
        //宣告並設定 終端機電腦記憶體的暫存物件『datasetNum』
        DataSet datasetNum = new DataSet();
        OleDbDataAdapter adapter;

        QueryResultForm qrForm = new QueryResultForm();

        public Main()
        {
            InitializeComponent();

            editStatus = 0;
            messageLabel.Text = "";

#if (NET45 || NET48)
            this.btnCrawl.Click += new System.EventHandler(this.BtnCrawl_Click);
            btnCrawl.Visible = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13 | SecurityProtocolType.Ssl3;
#else
            //may not work for https website
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Ssl3;
#endif
            ServicePointManager.DefaultConnectionLimit = 50;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                conn = new OleDbConnection(DB_CONN);
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
                string sql = "CREATE TABLE dailycash (日期 int, " +
                "一 INT, 二 INT, 三 INT, 四 INT, 五 INT)";
                OleDbCommand cmd = new OleDbCommand(sql, conn);
                cmd.ExecuteNonQuery();
            }
            catch (Exception e1)
            {
                Console.WriteLine(e1.Message);
            }

            refreshData();
            gotoRow(currentRow);
            updateDate();
        }

        private void Main_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (conn.State == ConnectionState.Open)
                conn.Close();
        }

        //更新資料庫內容到異動datagrid上
        private void refreshData()
        {
            datasetNum.Clear();

            //oledbdataadapter物件建立資料表查詢結果
            //	宣告並設定  查詢『num 資料表』字串
            string str = "SELECT * FROM dailycash ORDER BY 日期";
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
        private void gotoRow(int rowNumber)
        {
            objDV.ClearSelection();

            clearColor(objDV);

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
                btnFirst.Enabled = false;
                btnPrevious.Enabled = false;
            }
            else
            {
                btnFirst.Enabled = true;
                btnPrevious.Enabled = true;
            }

            if (rowNumber >= datasetNum.Tables["dailycash"].Rows.Count - 1)
            {
                btnNext.Enabled = false;
                btnLast.Enabled = false;
            }
            else
            {
                btnNext.Enabled = true;
                btnLast.Enabled = true;
            }
            objDV.Rows[rowNumber].Selected = true;
        }

        //update the newest date for every queries
        private void updateDate()
        {
            qDateTextBox2.Text = newestDate;
        }

        //清除cellcolor
        private void clearColor(DataGridView dgv)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
                for (int j = 0; j < dgv.Columns.Count; j++)
                    dgv.Rows[i].Cells[j].Style.BackColor = Color.Empty;
        }

        private void changeColor(DataGridView dgv, int num, Color color)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
                for (int j = 1; j < dgv.Columns.Count; j++)
                    if (Convert.ToInt32(dgv.Rows[i].Cells[j].Value) == num)
                        dgv.Rows[i].Cells[j].Style.BackColor = color;
                    else if (Convert.ToInt32(dgv.Rows[i].Cells[j].Value) % 10 == (num % 10))
                        dgv.Rows[i].Cells[j].Style.BackColor = Color.LightBlue;
            dgv.ClearSelection();
        }

        //第一筆按鈕
        private void FirstButton_Click(object sender, EventArgs e)
        {
            currentRow = 0;
            gotoRow(currentRow);
        }

        //上一筆按鈕
        private void PreButton_Click(object sender, EventArgs e)
        {
            currentRow -= 1;
            gotoRow(currentRow);
        }

        //下一筆按鈕
        private void NextButton_Click(object sender, EventArgs e)
        {
            currentRow += 1;
            gotoRow(currentRow);
        }

        //最後一筆按鈕
        private void LastButton_Click(object sender, EventArgs e)
        {
            currentRow = objDV.Rows.Count - 1;
            gotoRow(currentRow);
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
            btnDBOK.Enabled = true;
            btnDBCreate.Enabled = false;
            btnDBUpdate.Enabled = false;
            btnDBDelete.Enabled = false;
            btnFirst.Enabled = false;
            btnPrevious.Enabled = false;
            btnNext.Enabled = false;
            btnLast.Enabled = false;
        }

        //關閉編輯模式
        private void disableEdit()
        {
            dateTextBox.Enabled = false;
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            textBox3.Enabled = false;
            textBox4.Enabled = false;
            textBox5.Enabled = false;
            btnDBOK.Enabled = false;
            btnDBCreate.Enabled = true;
            btnDBUpdate.Enabled = true;
            btnDBDelete.Enabled = true;
            btnFirst.Enabled = true;
            btnPrevious.Enabled = true;
            btnNext.Enabled = true;
            btnLast.Enabled = true;
        }

#if (NET45 || NET48)
        private void BtnCrawl_Click(object sender, EventArgs e)
        {

            string urlAddress = "https://www.taiwanlottery.com.tw/Lotto/Dailycash/history.aspx";
            
            NameValueCollection nvc = new NameValueCollection();
            //nvc.Add("__EVENTTARGET", "");
            //nvc.Add("__EVENTARGUMENT", "");
            //nvc.Add("__LASTFOCUS", "");
            for (int nMonth = 1; nMonth <= 2; nMonth++)
            {
                nvc.Clear();;
                nvc.Add("__VIEWSTATE", "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");
                nvc.Add("__EVENTVALIDATION", "LMNhrZdgFJgWKAhvvqdec7uBDqZi8dalKYpYPq9PKuR5u7V/PIGU2iWvtlv2Iztubj5YOGGXfylzFp46lBcwU7euiQ7vjSwVNLuxfwuwRwwZUO1XZ3tIWhw1YkgQAy3oOsCoFkRWbiCntkAg9RFaZwpjRopfi8jNKtETmKFmYDAVrp5iYVIDOPGFwmweZTZR4tf/ixYzyn2/ZsMF39KcxhOrPKSQzUdQmxFOXECH4jxdN9yDgrB3hC8pQY11suKgdSjEu0aPgn+JecPRhxN2b20bmtQEuWOTOdzN8zFggRnOYehvDp89pSsIUOhkylfmSFLWjZSX+8qOY3R7A368KLDTWR63ecvDUSw1FLPXK7UayQE+EBtXKmuianYSxwImZO4lHkiqoEJN9iEG2CRFG2nud6586m/gyiOBU4SafzIdw16YVsHYVvYrWuqsI9SRnDPlI9KSzzNwuAwk/MmEmhj6XihRbViZaWW5b6RpnzgUdn7dgChHPzkOl6GU57IQd102g1PFX+xWNigQO6G49RfjAQNlNIAgPfQdqTkOAAr5qDuMNTNLMtyjfXhK5QlBxXI1n5Lq0qQmXsZ5nJPSqBBJOjdzxes4M0aSnfmZFPaZ7fDl7K7WEd5kQEfQ4Bn9oG7r4hA8fuv6/PE2njRvcEmNHBffNvFQ7k7Ynk4ap3Vz9npKOZ/Oc7jRtyobRS0gtCnA7MIhN3lTEkfIYtFC80uvfJz6ezt2fj2ZZcDiPgcrbc6AgW61dlnbAwikK/Bxvyz19M5nW2EKJaattxgAJ113h7RVPC+nzeTomQn0Ni5vcrn4NLXbd++DWywGhbDFEeBdSEAqSmFDrsbGyahwLbHSvko=");
                nvc.Add("__VIEWSTATEGENERATOR", "09BD3138");
                nvc.Add("D539Control_history1$DropDownList1", "5");
                nvc.Add("D539Control_history1$chk", "radYM");
                nvc.Add("D539Control_history1$dropYear", "109");
                nvc.Add("D539Control_history1$dropMonth", nMonth.ToString());
                nvc.Add("D539Control_history1$btnSubmit", "查詢");

                //Console.WriteLine(AssemblePostPayload(nvc));
 
                byte[] buff = Encoding.UTF8.GetBytes(AssemblePostPayload(nvc));
#if !_DEBUG
                HtmlWeb web = new HtmlWeb();

                HtmlWeb.PreRequestHandler handler = delegate (HttpWebRequest request)
                {
                    if (request.Method == "POST")
                    {
                        request.ContentType = "application/x-www-form-urlencoded";
                        request.GetRequestStream().Write(buff, 0, buff.Length);
                    }

                    return true;
                };

                web.PreRequest += handler;
                var doc = web.Load(urlAddress, "POST");
                //Console.WriteLine(doc.Text);
                web.PreRequest -= handler;
                var children = doc.DocumentNode.SelectNodes("//span[contains(@id,'D539Control_history1_dlQuery_D539_DDate')] | //span[contains(@id,'D539Control_history1_dlQuery_No')]");

                for (int i = 0; i < children.Count; i += 6)
                {
                    string strDate = System.Text.RegularExpressions.Regex.Replace(children[i].InnerText, "/", "");
                    string sql = string.Format("INSERT INTO dailycash (日期, 一, 二, 三, 四, 五) VALUES ({0},{1}", strDate, children[i + 1].InnerText);
                    for (int j = 2; j <= 5; j++)
                        sql += "," + children[i + j].InnerText;
                    sql += ")";
                    //Console.WriteLine(sql);

                    var obj = new OleDbCommand(string.Format("SELECT * FROM dailycash WHERE 日期={0}", strDate), conn).ExecuteScalar();
                    if (obj == null)
                    {
                        OleDbCommand cmd = new OleDbCommand(sql, conn);
                        //執行資料庫指令OleDbCommand
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            refreshData();
            currentRow = 0;
            gotoRow(currentRow);
            updateDate();
#else
                var request = (HttpWebRequest)WebRequest.Create(urlAddress);
            var str = "";

            request.ServicePoint.Expect100Continue = false;
            request.ServicePoint.UseNagleAlgorithm = false;
            request.AllowAutoRedirect = false;
            request.AllowWriteStreamBuffering = false;
            request.Headers.Add(HttpRequestHeader.AcceptEncoding, "gzip,deflate, br");
            request.Accept = "*/*";
            request.Headers.Add(HttpRequestHeader.AcceptLanguage, "zh-TW,zh-CN,zh;q=0.8,en-US;q=0.7,en;q=6");
            //request.ContentType = "application/x-www-form-urlencoded";
            request.UserAgent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.106 Safari/537.36";
            request.Timeout = 5000;
            request.Method = "POST";
            request.ContentLength = buff.Length;
            request.GetRequestStream().Write(buff, 0, buff.Length);

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
                }
            }
#endif
        }

        protected void AppendParameter(StringBuilder sb, string name, string value)
        {
            string encodedValue = System.Net.WebUtility.UrlEncode(value);
            sb.AppendFormat("{0}={1}&", name, encodedValue);
        }

        private string AssemblePostPayload(NameValueCollection fv)
        {
            StringBuilder sb = new StringBuilder();
            foreach (String key in fv.AllKeys)
            {
                sb.Append("&" + Uri.EscapeDataString(key) + "=" + Uri.EscapeDataString(fv.Get(key)));
            }
            return sb.ToString().Substring(1);
        }
#endif

        void btnDBCreate_Click(object sender, EventArgs e)
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

        private void btnDBUpdate_Click(object sender, EventArgs e)
        {
            EnableEdit();
            editStatus = 2;
            btnDBCreate.Enabled = false;
            btnDBDelete.Enabled = false;
        }

        private void btnDBDelete_Click(object sender, EventArgs e)
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

                refreshData();
                currentRow = 0;
                gotoRow(currentRow);
            }
            updateDate();
        }

        private void btnDBOK_Click(object sender, EventArgs e)
        {
            DialogResult answer;

            if (editStatus == 1)
            {
                answer = MessageBox.Show("您確定要新增此筆資料嗎？", "新增資料", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

                if (answer == DialogResult.Yes)
                {
                    //新增記錄到資料庫內
                    //	設定新增記錄的 SQL語法 及資料庫執行指令OleDbCommand
                    int nDate = Int32.Parse(dateTextBox.Text);
                    var obj = new OleDbCommand(string.Format("SELECT * FROM dailycash WHERE 日期={0}", nDate), conn).ExecuteScalar();
                    if (obj == null)
                    {
                        string str = string.Format("INSERT INTO dailycash (日期,一,二,三,四,五) VALUES ({0},{1},{2},{3},{4},{5})",
                            Int32.Parse(dateTextBox.Text),
                            Int32.Parse(textBox1.Text),
                            Int32.Parse(textBox2.Text),
                            Int32.Parse(textBox3.Text),
                            Int32.Parse(textBox4.Text),
                            Int32.Parse(textBox5.Text));
                        Console.WriteLine(str);

                        OleDbCommand cmd = new OleDbCommand(str, conn);

                        //執行資料庫指令OleDbCommand
                        cmd.ExecuteNonQuery();
                    }

                    refreshData();
                    currentRow = 0;
                    gotoRow(currentRow);
                }
            }
            else if (editStatus == 2)
            {
                answer = MessageBox.Show("您確定要修改此筆資料嗎？", "修改資料", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

                if (answer == DialogResult.Yes)
                {
                    //修改資料庫內的記錄
                    //	設定修改記錄的  SQL語法及資料庫執行指令OleDbCommand
                    string str = "Update dailycash set 日期 = " + Int32.Parse(dateTextBox.Text) + ",一=" + Int32.Parse(textBox1.Text) + ",二=" + Int32.Parse(textBox2.Text) + ",三=" + Int32.Parse(textBox3.Text) + ",四=" + Int32.Parse(textBox4.Text) + ",五=" + Int32.Parse(textBox5.Text) + " WHERE 日期= " + Int32.Parse(dateTextBox.Text) + "";

                    OleDbCommand cmd = new OleDbCommand(str, conn);

                    //執行資料庫指令OleDbCommand
                    cmd.ExecuteNonQuery();

                    refreshData();
                    gotoRow(currentRow);
                }
            }
            disableEdit();
            updateDate();
        }

        private void objDV_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                currentRow = e.RowIndex;
                gotoRow(currentRow);
            }
        }

        private void objDV_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex != -1 && e.ColumnIndex != 0)
            {
                clearColor(objDV);
                objDV.CurrentCell.Style.BackColor = Color.Yellow;
                changeColor(objDV, Convert.ToInt32(objDV.CurrentCell.Value), Color.Yellow);
            }
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            //oledbdataadapter物件建立資料表查詢結果
            //	宣告並設定  查詢『dailycash 資料表』字串
            string str = "SELECT * FROM dailycash WHERE 日期 BETWEEN " + qDateTextBox1.Text + " AND " + qDateTextBox2.Text + " ORDER BY 日期";

            //	宣告並設定  資料表查詢物件『adapter』
            adapter = new OleDbDataAdapter(str, conn);

            DataTable table = new DataTable();
            DataTable oriTable = new DataTable();
            adapter.Fill(table);
            adapter.Fill(oriTable);

            if (table.Rows.Count != 0)
            {
                //宣告並判斷後幾期變數,預設為0
                int rangeNum = 0;
                if (rangeTextBox.Text == "")
                {
                    rangeTextBox.Text = rangeNum.ToString();
                }
                else
                    rangeNum = Int32.Parse(rangeTextBox.Text);

                bool[] status = new bool[6];
                int[] inputNum = new int[5];

                //判斷是否輸入錯誤(未輸入所有查詢值)
                if ((textBox_1.Text == "") && (textBox_2.Text == "") && (textBox_3.Text == "") && (textBox_4.Text == "") && (textBox_5.Text == ""))
                    status[0] = false;
                else
                    status[0] = true;

                //判斷輸入值
                //1
                if (textBox_1.Text == "")
                    status[1] = true;
                else
                    inputNum[0] = Int32.Parse(textBox_1.Text);
                //2
                if (textBox_2.Text == "")
                    status[2] = true;
                else
                    inputNum[1] = Int32.Parse(textBox_2.Text);
                //3
                if (textBox_3.Text == "")
                    status[3] = true;
                else
                    inputNum[2] = Int32.Parse(textBox_3.Text);
                //4
                if (textBox_4.Text == "")
                    status[4] = true;
                else
                    inputNum[3] = Int32.Parse(textBox_4.Text);
                //5
                if (textBox_5.Text == "")
                    status[5] = true;
                else
                    inputNum[4] = Int32.Parse(textBox_5.Text);

                if (status[0])
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //符合輸入數值的資料
                        if ((status[1] || (System.Convert.ToInt32(table.Rows[i][1]) == inputNum[0])) && (status[2] || (System.Convert.ToInt32(table.Rows[i][2]) == inputNum[1])) && (status[3] || (System.Convert.ToInt32(table.Rows[i][3]) == inputNum[2])) && (status[4] || (System.Convert.ToInt32(table.Rows[i][4]) == inputNum[3])) && (status[5] || (System.Convert.ToInt32(table.Rows[i][5]) == inputNum[4])))
                        {
                            //下j期迴圈
                            for (int j = 1; j <= rangeNum; j++)
                            {
                                //當下j期存在於表格內
                                if (i + j < table.Rows.Count)
                                {
                                    if ((status[1] || (System.Convert.ToInt32(table.Rows[i + j][1]) == inputNum[0])) && (status[2] || (System.Convert.ToInt32(table.Rows[i + j][2]) == inputNum[1])) && (status[3] || (System.Convert.ToInt32(table.Rows[i + j][3]) == inputNum[2])) && (status[4] || (System.Convert.ToInt32(table.Rows[i + j][4]) == inputNum[3])) && (status[5] || (System.Convert.ToInt32(table.Rows[i + j][5]) == inputNum[4])))
                                    {
                                        i = i + j;
                                        break;
                                    }
                                    else
                                        continue;
                                }
                            }
                            i = i + rangeNum;
                        }
                        else
                            table.Rows[i].Delete();
                    }
                    table.AcceptChanges();

                    objDV1.DataSource = table;

                    //20100704 refined
                    qrForm.totalTable.Merge(table);
                    qrForm.refresh();
                    qrForm.Show();

                    objDV1.ClearSelection();
                    int[] itemArr = findIndexTable(oriTable, table);
                    sameColorNum(objDV1, itemArr, oriTable, table);
                }
                else
                    MessageBox.Show("請輸入欲查詢數字！");
            }
            else
                MessageBox.Show("查無此範圍資料...");
        }

        private int[] findIndexTable(DataTable table1, DataTable table2)
        {
            int[] arr = new int[table2.Rows.Count];

            for (int j = 0; j < table2.Rows.Count; j++)
                for (int i = 0; i < table1.Rows.Count; i++)
                {
                    if (table1.Rows[i]["日期"].Equals(table2.Rows[j]["日期"]))
                        arr[j] = i;
                }
            return arr;
        }

        private void sameColorNum(DataGridView dgv, int[] array, DataTable oTable, DataTable mTable)
        {
            int[] arr = sort(array);
            if (arr.Length != 0)
            {
                //counter表示查詢結果有幾組
                int counter = 1;
                for (int i = 0; i < arr.Length; i++)
                    if ((i + 1) == arr.Length)
                        break;
                    else if (arr[i] != (arr[i + 1] - 1))
                        counter++;

                int[,] temp = new int[counter, 2];

                for (int i = 0; i < counter; i++)
                {
                    if (i == 0)
                    {
                        temp[0, 0] = arr[0];
                        temp[0, 1] = 1;
                    }
                    for (int j = 1; j < arr.Length; j++)
                        if (arr[j] == (arr[j - 1] + 1))
                            temp[i, 1]++;
                        else
                        {
                            i++;
                            temp[i, 0] = arr[j];
                            temp[i, 1] = 1;
                        }
                }

                for (int i = 0; i < counter; i++)
                    if (i % 2 == 1)
                    {
                        for (int j = 0; j < temp[i, 1] - 1; j++)
                            paint(dgv, oTable.Rows[temp[i, 0] + j], oTable.Rows[temp[i, 0] + j + 1], Color.LightPink);
                    }
                    else
                    {
                        for (int k = 0; k < temp[i, 1]; k++)
                            paintBack(dgv, oTable.Rows[temp[i, 0] + k], Color.LightGray);

                        for (int j = 0; j < temp[i, 1] - 1; j++)
                            paint(dgv, oTable.Rows[temp[i, 0] + j], oTable.Rows[temp[i, 0] + j + 1], Color.GreenYellow);
                    }
            }
        }

        private int[] sort(int[] arr)
        {
            for (int i = 0; i < arr.Length - 1; i++)
                for (int j = i + 1; j < arr.Length; j++)
                    if (arr[i] > arr[j])
                    {
                        int temp = arr[i];
                        arr[i] = arr[j];
                        arr[j] = temp;
                    }
            return arr;
        }

        private void paint(DataGridView dgv, DataRow row1, DataRow row2, Color color)
        {
            int m = -1, n = -1;
            for (int i = 0; i < dgv.Rows.Count; i++)
                if (dgv.Rows[i].Cells["日期"].Value.Equals(row1["日期"]))
                    m = i;
                else if (dgv.Rows[i].Cells["日期"].Value.Equals(row2["日期"]))
                    n = i;

            for (int i = 1; i <= 5; i++)
                for (int j = 1; j <= 5; j++)
                    if (dgv.Rows[m].Cells[i].Value.Equals(dgv.Rows[n].Cells[j].Value))
                    {
                        dgv.Rows[m].Cells[i].Style.BackColor = color;
                        dgv.Rows[n].Cells[j].Style.BackColor = color;
                        break;

                    }
        }
        //將傳入之row塗上color
        private void paintBack(DataGridView dgv, DataRow row, Color color)
        {
            int m = -1;
            for (int i = 0; i < dgv.Rows.Count; i++)
                if (dgv.Rows[i].Cells["日期"].Value.Equals(row["日期"]))
                {
                    m = i;
                    break;
                }
            for (int i = 0; i <= 5; i++)
                dgv.Rows[m].Cells[i].Style.BackColor = color;
        }

        private void objDV1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            clearColor(objDV1);
            objDV1.CurrentCell.Style.BackColor = Color.Yellow;
            changeColor(objDV1, Convert.ToInt32(objDV1.CurrentCell.Value), Color.Yellow);
        }
    }
}
