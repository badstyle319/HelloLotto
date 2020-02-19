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
            btnCrawl.Visible = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13 | SecurityProtocolType.Ssl3;
#else
            //may not work for https website
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Ssl3;
#endif
            //ServicePointManager.DefaultConnectionLimit = 50;
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

        private void BtnCrawl_Click(object sender, EventArgs e)
        {
            string urlAddress = "https://www.taiwanlottery.com.tw/Lotto/Dailycash/history.aspx";

#if !_DEBUG
            List<HtmlNode> section = new List<HtmlNode>();

            NameValueCollection postData = new NameValueCollection();
            postData.Add("D539Control_history1$DropDownList1", "5");
            postData.Add("D539Control_history1$chk", "radYM");
            postData.Add("D539Control_history1$dropYear", "109");
            postData.Add("D539Control_history1$dropMonth", "1");
            postData.Add("D539Control_history1$btnSubmit", "查詢");

            HtmlWeb web = new HtmlWeb();

            HtmlWeb.PreRequestHandler handler = delegate (HttpWebRequest request)
            {
                request.ServicePoint.Expect100Continue = false;
                request.AllowAutoRedirect = false;
                request.CookieContainer = new CookieContainer();
                string payLoad = AssemblePostPayload(postData);
                byte[] buff = Encoding.UTF8.GetBytes(payLoad.ToCharArray());
                request.ContentLength = buff.Length;
                request.ContentType = "application/x-www-form-urlencoded";
                request.GetRequestStream().Write(buff, 0, buff.Length);
                Console.WriteLine(buff.Length);
                return true;
            };

            web.PreRequest += handler;
            var doc = web.Load(urlAddress, "POST");
            web.PreRequest -= handler;
            var children = doc.DocumentNode.SelectNodes("//span[contains(@id,'D539Control_history1_dlQuery_D539_DDate')] | //span[contains(@id,'D539Control_history1_dlQuery_SNo')]");

            for (int i = 0; i < children.Count; i += 6)
            {
                string strDate = System.Text.RegularExpressions.Regex.Replace(children[i].InnerText, "/", "");
                string sql = string.Format("INSERT INTO dailycash (日期, 一, 二, 三, 四, 五) VALUES ({0},{1}", strDate, children[i + 1].InnerText);
                for (int j = 2; j <= 5; j++)
                    sql += "," + children[i + j].InnerText;
                sql += ")";
                Console.WriteLine(sql);

                var obj = new OleDbCommand(string.Format("SELECT * FROM dailycash WHERE 日期={0}", strDate), conn).ExecuteScalar();
                if (obj == null)
                {
                    OleDbCommand cmd = new OleDbCommand(sql, conn);
                    //執行資料庫指令OleDbCommand
                    cmd.ExecuteNonQuery();
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
                }
            }
#endif
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

        private void btnDBCreate_Click(object sender, EventArgs e)
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
