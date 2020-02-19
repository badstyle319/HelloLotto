using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace _0726
{
    public partial class Form1 : Form
    {
        //記錄目前所選資料索引
        static int currentRow = 0;

        //connection連結到資料庫
        //	宣告並設定  連接字串
        static string strConn = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=LT.mdb";
        //	宣告並設定  連接物件conn
        OleDbConnection conn = new OleDbConnection(strConn);
        //dataset11物件置於暫時記憶體，以存放查詢結果
        //	宣告並設定 終端機電腦記憶體的暫存物件『datasetNum』
        DataSet datasetNum = new DataSet();
        OleDbDataAdapter adapter;
        int editStatus;//1表示新增資料,2表示修改資料
        //記錄最新一筆資料之日期
        string newestDate = "";
        //舊查詢中之統計數字陣列
        int[] totalNum = new int[49];
        QueryResultForm qrForm = new QueryResultForm();

        public Form1()
        {
            InitializeComponent();
            editStatus = 0;
            //test();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            refreshData();
            gotoRow(currentRow);
            updateDate();
        }

        //update the newest date for every queries
        private void updateDate()
        {
            qDateTextBox2.Text = newestDate;
            pDateTextBox2.Text = newestDate;
            sDateTextBox2.Text = newestDate;
            q2DateTextBox2.Text = newestDate;
            aQueryDateTextBox2.Text = newestDate;
        }

        //更新資料庫內容到異動datagrid上
        private void refreshData()
        {
            datasetNum.Clear();
            //	進行連結資料庫
            conn.Open();
            //oledbdataadapter物件建立資料表查詢結果
            //	宣告並設定  查詢『num 資料表』字串
            string str = "select * from num order by 日期";
            //	宣告並設定  資料表查詢物件『adapter』
            adapter = new OleDbDataAdapter(str, conn);

            //	將伺服器資料庫的查詢結果（adapter）存放並填滿到終端機的暫存物件（datasetNum）上的表格"num"
            adapter.Fill(datasetNum, "num");

            //在datagrid來顯示dataset上的資料
            objDV.DataSource = datasetNum.Tables["num"];
            newestDate = datasetNum.Tables["num"].Rows[datasetNum.Tables["num"].Rows.Count - 1]["日期"].ToString();
            //關閉連線
            conn.Close();

            objDV.ClearSelection();
        }

        //跳到第N筆資料方法
        private void gotoRow(int rowNumber)
        {
            objDV.ClearSelection();

            clearColor(objDV);

            dateTextBox.Text = datasetNum.Tables["num"].Rows[rowNumber]["日期"].ToString();
            textBox1.Text = datasetNum.Tables["num"].Rows[rowNumber]["一"].ToString();
            textBox2.Text = datasetNum.Tables["num"].Rows[rowNumber]["二"].ToString();
            textBox3.Text = datasetNum.Tables["num"].Rows[rowNumber]["三"].ToString();
            textBox4.Text = datasetNum.Tables["num"].Rows[rowNumber]["四"].ToString();
            textBox5.Text = datasetNum.Tables["num"].Rows[rowNumber]["五"].ToString();
            textBox6.Text = datasetNum.Tables["num"].Rows[rowNumber]["六"].ToString();
            textBox7.Text = datasetNum.Tables["num"].Rows[rowNumber]["特"].ToString();
            messageLabel.Text = "第 " + (rowNumber + 1) + " 筆，共有 " + datasetNum.Tables["num"].Rows.Count + " 筆資料。";

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

            if (rowNumber >= datasetNum.Tables["num"].Rows.Count - 1)
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

        //第一筆按鈕
        private void firstButton_Click(object sender, EventArgs e)
        {
            currentRow = 0;
            gotoRow(currentRow);
        }

        //上一筆按鈕
        private void preButton_Click(object sender, EventArgs e)
        {
            currentRow -= 1;
            gotoRow(currentRow);
        }

        //下一筆按鈕
        private void nextButton_Click(object sender, EventArgs e)
        {
            currentRow += 1;
            gotoRow(currentRow);
        }

        //最後一筆按鈕
        private void lastButton_Click(object sender, EventArgs e)
        {
            currentRow = objDV.Rows.Count - 1;
            gotoRow(currentRow);
        }

        //啟動編輯模式
        private void enableEdit()
        {
            dateTextBox.Enabled = true;
            textBox1.Enabled = true;
            textBox2.Enabled = true;
            textBox3.Enabled = true;
            textBox4.Enabled = true;
            textBox5.Enabled = true;
            textBox6.Enabled = true;
            textBox7.Enabled = true;
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
        private void disableEdit()
        {
            dateTextBox.Enabled = false;
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            textBox3.Enabled = false;
            textBox4.Enabled = false;
            textBox5.Enabled = false;
            textBox6.Enabled = false;
            textBox7.Enabled = false;
            okButton.Enabled = false;
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            firstButton.Enabled = true;
            preButton.Enabled = true;
            nextButton.Enabled = true;
            lastButton.Enabled = true;
        }

        //新增資料按鈕
        private void button1_Click(object sender, EventArgs e)
        {
            enableEdit();
            dateTextBox.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            editStatus = 1;
            dateTextBox.Focus();
        }

        //修改資料按鈕
        private void button2_Click(object sender, EventArgs e)
        {
            enableEdit();
            editStatus = 2;
            button1.Enabled = false;
            button3.Enabled = false;
        }

        //刪除資料按鈕
        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult answer = MessageBox.Show("您確定要刪除此筆資料嗎？", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

            if (answer == DialogResult.Yes)
            {
                conn.Open();

                //刪除資料庫內的記錄
                //	設定刪除記錄的 SQL語法 及資料庫執行指令OleDbCommand
                string str = "delete * from num where 日期=" + Int32.Parse(dateTextBox.Text) + "";
                OleDbCommand cmd = new OleDbCommand(str, conn);

                //執行資料庫指令OleDbCommand
                cmd.ExecuteNonQuery();

                //關閉資料庫連接
                conn.Close();

                refreshData();
                currentRow = 0;
                gotoRow(currentRow);
            }
            updateDate();
        }

        //異動dataGridView事件處理
        private void objDV_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)
            {
                currentRow = e.RowIndex;
                gotoRow(currentRow);
            }
        }


        //查詢按鈕
        private void queryButton_Click(object sender, EventArgs e)
        {
            //	進行連結資料庫
            conn.Open();
            //oledbdataadapter物件建立資料表查詢結果
            //	宣告並設定  查詢『num 資料表』字串
            string str = "select * from num where 日期 between " + qDateTextBox1.Text + " and " + qDateTextBox2.Text + " order by 日期";

            //	宣告並設定  資料表查詢物件『adapter』
            adapter = new OleDbDataAdapter(str, conn);

            DataTable table = new DataTable();
            DataTable oriTable = new DataTable();
            adapter.Fill(table);
            adapter.Fill(oriTable);

            //關閉連線
            conn.Close();

            if (table.Rows.Count != 0)
            {

                //宣告並判斷後幾期變數,預設為0
                int rangeNum = 0;
                if (rangeTextBox.Text == "")
                {
                    rangeNum = 0;
                    rangeTextBox.Text = rangeNum.ToString();
                }
                else
                    rangeNum = Int32.Parse(rangeTextBox.Text);

                bool[] status = new bool[8];
                int[] inputNum = new int[7];

                //判斷是否輸入錯誤(未輸入所有查詢值)
                if ((textBox_1.Text == "") && (textBox_2.Text == "") && (textBox_3.Text == "") && (textBox_4.Text == "") && (textBox_5.Text == "") && (textBox_6.Text == "") && (textBox_7.Text == ""))
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
                //6
                if (textBox_6.Text == "")
                    status[6] = true;
                else
                    inputNum[5] = Int32.Parse(textBox_6.Text);
                //7
                if (textBox_7.Text == "")
                    status[7] = true;
                else
                    inputNum[6] = Int32.Parse(textBox_7.Text);

                if (status[0])
                {
                    /**
                    for (int i = table.Rows.Count - 1; i - rangeNum >= 0; i--)
                    {
                        if ((status[1] || (System.Convert.ToInt32(table.Rows[i - rangeNum][1]) == inputNum[0])) && (status[2] || (System.Convert.ToInt32(table.Rows[i - rangeNum][2]) == inputNum[1])) && (status[3] || (System.Convert.ToInt32(table.Rows[i - rangeNum][3]) == inputNum[2])) && (status[4] || (System.Convert.ToInt32(table.Rows[i - rangeNum][4]) == inputNum[3])) && (status[5] || (System.Convert.ToInt32(table.Rows[i - rangeNum][5]) == inputNum[4])) && (status[6] || (System.Convert.ToInt32(table.Rows[i - rangeNum][6]) == inputNum[5])) && (status[7] || (System.Convert.ToInt32(table.Rows[i - rangeNum][7]) == inputNum[6])))
                            continue;
                        else
                            table.Rows[i].Delete();
                    }
                    for (int i = 0; i < rangeNum; i++)
                        table.Rows[i].Delete();
                    **/
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        //符合輸入數值的資料
                        if ((status[1] || (System.Convert.ToInt32(table.Rows[i][1]) == inputNum[0])) && (status[2] || (System.Convert.ToInt32(table.Rows[i][2]) == inputNum[1])) && (status[3] || (System.Convert.ToInt32(table.Rows[i][3]) == inputNum[2])) && (status[4] || (System.Convert.ToInt32(table.Rows[i][4]) == inputNum[3])) && (status[5] || (System.Convert.ToInt32(table.Rows[i][5]) == inputNum[4])) && (status[6] || (System.Convert.ToInt32(table.Rows[i][6]) == inputNum[5])) && (status[7] || (System.Convert.ToInt32(table.Rows[i][7]) == inputNum[6])))
                        {
                            //下j期迴圈
                            for (int j = 1; j <= rangeNum; j++)
                            {
                                //當下j期存在於表格內
                                if (i + j < table.Rows.Count)
                                {
                                    if ((status[1] || (System.Convert.ToInt32(table.Rows[i + j][1]) == inputNum[0])) && (status[2] || (System.Convert.ToInt32(table.Rows[i + j][2]) == inputNum[1])) && (status[3] || (System.Convert.ToInt32(table.Rows[i + j][3]) == inputNum[2])) && (status[4] || (System.Convert.ToInt32(table.Rows[i + j][4]) == inputNum[3])) && (status[5] || (System.Convert.ToInt32(table.Rows[i + j][5]) == inputNum[4])) && (status[6] || (System.Convert.ToInt32(table.Rows[i + j][6]) == inputNum[5])) && (status[7] || (System.Convert.ToInt32(table.Rows[i + j][7]) == inputNum[6])))
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

        //統計按鈕
        /**
        private void button4_Click(object sender, EventArgs e)
        {
            
            //	進行連結資料庫
            conn.Open();
            //oledbdataadapter物件建立資料表查詢結果
            //	宣告並設定  查詢『num 資料表』字串
            string str = "select 一,二,三,四,五,六,特 from num where 日期 between "+dateTextBox1.Text+" and "+dateTextBox2.Text+" order by 日期";

            //	宣告並設定  資料表查詢物件『adapter』
            adapter = new OleDbDataAdapter(str, conn);

            DataTable table = new DataTable();
            DataTable result1 = new DataTable();
            DataTable result2 = new DataTable();
            DataTable pTable = new DataTable();
            DataTable p1 = new DataTable();
            DataTable p2 = new DataTable();

            //	將伺服器資料庫的查詢結果（adapter）存放到table上
            adapter.Fill(table);

            str = "select * from result1";
            adapter = new OleDbDataAdapter(str, conn);
            adapter.Fill(result1);

            str = "select * from result2";
            adapter = new OleDbDataAdapter(str, conn);
            adapter.Fill(result2);

            str = "select * from num where 日期 between " + dateTextBox1.Text + " and " + dateTextBox2.Text + " order by 日期";
            adapter = new OleDbDataAdapter(str, conn);
            adapter.Fill(pTable);

            str = "select * from p1";
            adapter = new OleDbDataAdapter(str, conn);
            adapter.Fill(p1);

            str = "select * from p2";
            adapter = new OleDbDataAdapter(str, conn);
            adapter.Fill(p2);

            Form2 pForm = new Form2();

            try
            {
                for (int i = 0; i < pTable.Rows.Count; i++)
                {
                    p1.Rows[i][0] = pTable.Rows[i][0];
                    p2.Rows[i][0] = pTable.Rows[i][0];    
                }

                pForm.pGridView1.DataSource = p1;
                pForm.pGridView2.DataSource = p2;

                pForm.Show();

                for (int i = 0; i < pTable.Rows.Count; i++)
                    for (int j = 1; j < 8; j++)
                    {
                        string temp = pTable.Rows[i][j].ToString();
                        if (Convert.ToInt16(temp) <= 25)
                            pForm.pGridView1.Rows[i].Cells[temp].Style.BackColor = Color.Black;
                        else
                            pForm.pGridView2.Rows[i].Cells[temp].Style.BackColor = Color.Black;
                    }
            }
            catch(Exception)
            {
                MessageBox.Show("指定範圍筆數超過一百筆..");
            }

            //關閉連線
            conn.Close();

            if (table.Rows.Count != 0)
            {
                //記錄號碼與尾數出現次數
                int[] count = new int[49];
                int[] tail = new int[10];

                for (int i = 0; i < table.Rows.Count; i++)
                    for (int j = 0; j < table.Columns.Count; j++)
                    {
                        int temp = System.Convert.ToInt16(table.Rows[i][j]);
                        count[temp - 1]++;
                        tail[temp % 10]++;
                    }

                int[] topFiveTail = findTopNIndex(tail, 5);
                int[] topTenNum = findTopNIndex(count, 10);

                //result1 process
                for (int i = 0; i < result1.Rows.Count; i++)
                {
                    result1.Rows[i]["號碼"] = topTenNum[i]+1;
                    result1.Rows[i]["次數"] = count[topTenNum[i]];
                    result1.Rows[i]["機率"] = count[topTenNum[i]] / 49.0 / 7.0;
                }


                //result2 process
                for (int i = 0; i < result2.Rows.Count; i++)
                {
                    result2.Rows[i]["尾數"] = topFiveTail[i];
                    int[] temp = new int[5];

                    for (int j = 0; j < 5; j++)
                        if (topFiveTail[i] == 0)
                        {
                            if (j == 0)
                                temp[j] = -1;
                            else
                                temp[j] = count[topFiveTail[i] - 1 + 10 * j];
                        }
                        else
                            temp[j] = count[topFiveTail[i] - 1 + 10 * j];

                    int[] top3 = findTopNIndex(temp, 3);
                    result2.Rows[i]["號碼一"] = 10 * top3[0] + Convert.ToInt16(result2.Rows[i]["尾數"]);
                    result2.Rows[i]["號碼二"] = 10 * top3[1] + Convert.ToInt16(result2.Rows[i]["尾數"]);
                    result2.Rows[i]["號碼三"] = 10 * top3[2] + Convert.ToInt16(result2.Rows[i]["尾數"]);
                }

                dataGridView1.DataSource = result1;
                dataGridView2.DataSource = result2;
            }
            else
                MessageBox.Show("查無此範圍資料...");
        }**/

        //編輯確定按鈕,editStatus=1表示新增.2表示修改
        private void okButton_Click(object sender, EventArgs e)
        {
            DialogResult answer;

            if (editStatus == 1)
            {
                answer = MessageBox.Show("您確定要新增此筆資料嗎？", "新增資料", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

                if (answer == DialogResult.Yes)
                {
                    conn.Open();

                    //新增記錄到資料庫內
                    //	設定新增記錄的 SQL語法 及資料庫執行指令OleDbCommand
                    var obj = new OleDbCommand(string.Format("SELECT * FROM num WHERE 日期={0}", Int32.Parse(dateTextBox.Text)), conn).ExecuteScalar();
                    if (obj == null)
                    {
                        string str = "INSERT INTO num(日期,一,二,三,四,五,六,特)VALUES(" + Int32.Parse(dateTextBox.Text) + "," + Int32.Parse(textBox1.Text) + "," + Int32.Parse(textBox2.Text) + "," + Int32.Parse(textBox3.Text) + "," + Int32.Parse(textBox4.Text) + "," + Int32.Parse(textBox5.Text) + "," + Int32.Parse(textBox6.Text) + "," + Int32.Parse(textBox7.Text) + ")";

                        OleDbCommand cmd = new OleDbCommand(str, conn);

                        //執行資料庫指令OleDbCommand
                        cmd.ExecuteNonQuery();
                    }
                    //關閉資料庫連接
                    conn.Close();

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
                    conn.Open();

                    //修改資料庫內的記錄
                    //	設定修改記錄的  SQL語法及資料庫執行指令OleDbCommand
                    string str = "Update num set 日期 = " + Int32.Parse(dateTextBox.Text) + ",一=" + Int32.Parse(textBox1.Text) + ",二=" + Int32.Parse(textBox2.Text) + ",三=" + Int32.Parse(textBox3.Text) + ",四=" + Int32.Parse(textBox4.Text) + ",五=" + Int32.Parse(textBox5.Text) + ",六=" + Int32.Parse(textBox6.Text) + ",特=" + Int32.Parse(textBox7.Text) + " where 日期= " + Int32.Parse(dateTextBox.Text) + "";

                    OleDbCommand cmd = new OleDbCommand(str, conn);

                    //執行資料庫指令OleDbCommand
                    cmd.ExecuteNonQuery();

                    //關閉資料庫連接
                    conn.Close();

                    refreshData();
                    gotoRow(currentRow);
                }
            }
            disableEdit();
            updateDate();
        }

        private int[] findTopNIndex(int[] num, int n)
        {
            int[] max = new int[n];
            int[] arr = new int[num.Length];
            num.CopyTo(arr, 0);

            for (int j = 0; j < n; j++)
            {
                for (int i = 0; i < num.Length; i++)
                {
                    int temp = arr.Max();

                    if (arr[i] == temp)
                    {
                        max[j] = i;
                        arr[i] = 0;
                        break;
                    }
                }
            }
            return max;
        }

        //判斷傳入之datarow是否含有數字陣列之每個值(true)或其中之一即可(false)
        private bool containNums(DataRow row, int[] num, bool status)
        {
            if (status)
            {
                int counter = 0;
                for (int i = 1; i <= 7; i++)
                    if (num.Contains(Convert.ToInt32(row[i])))
                        counter++;
                if (counter == num.Length)
                    return true;
                else
                    return false;
            }
            else
            {
                for (int i = 1; i <= 7; i++)
                    if (num.Contains(Convert.ToInt32(row[i])))
                        return true;
                return false;
            }
        }
        //刪去table中不含所有陣列值的rows
        private void tableFilter(DataTable table, int[] num, bool status)
        {
            for (int i = 0; i < table.Rows.Count; i++)
                if (!containNums(table.Rows[i], num, status))
                {
                    table.Rows[i].Delete();

                }
            table.AcceptChanges();
        }
        //改變cellcolor
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
        //清除cellcolor
        private void clearColor(DataGridView dgv)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
                for (int j = 0; j < dgv.Columns.Count; j++)
                    dgv.Rows[i].Cells[j].Style.BackColor = Color.Empty;
        }

        private void objDV1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            clearColor(objDV1);
            objDV1.CurrentCell.Style.BackColor = Color.Yellow;
            changeColor(objDV1, Convert.ToInt32(objDV1.CurrentCell.Value), Color.Yellow);
        }

        enum numColumns : byte
        {
            一 = 1,
            二 = 2,
            三 = 3,
            四 = 4,
            五 = 5,
            六 = 6,
            特 = 7
        }

        //連莊號碼process
        private void sButton_Click(object sender, EventArgs e)
        {
            bool status = true;
            int num = 0;
            numColumns input = new numColumns();

            if (sLocTextBox.Text != "")
            {
                num = Convert.ToInt32(sLocTextBox.Text);
                if (num < 1 || num > 7)
                    status = false;
                else
                    input = (numColumns)num;
            }
            else
                status = false;

            if (sNumTextBox.Text == "")
                status = false;

            if (status)
            {
                //numColumns test = (numColumns)num;
                //	進行連結資料庫
                conn.Open();

                //oledbdataadapter物件建立資料表查詢結果
                //	宣告並設定  查詢『num 資料表』字串
                string str = "select 日期,一,二,三,四,五,六,特 from num where "
                    + input + "= " + sNumTextBox.Text//Convert.ToInt32(sNumTextBox.Text) 
                    + " and 日期 between " + sDateTextBox1.Text
                    + " and " + sDateTextBox2.Text + " order by 日期";

                //	宣告並設定  資料表查詢物件『adapter』
                adapter = new OleDbDataAdapter(str, conn);

                DataTable matchTable = new DataTable();
                DataTable oriTable = new DataTable();
                //	將伺服器資料庫的查詢結果（adapter）存放到table上
                adapter.Fill(matchTable);

                adapter = new OleDbDataAdapter("select * from num where 日期 between " + sDateTextBox1.Text + " and " + sDateTextBox2.Text + " order by 日期", conn);
                adapter.Fill(oriTable);
                //關閉連線
                conn.Close();

                filterTable(oriTable, matchTable, 1);
                sDataGridView.DataSource = matchTable;
                sDataGridView.ClearSelection();

                //20100704 refined
                qrForm.totalTable.Merge(matchTable);
                qrForm.refresh();
                qrForm.Show();

                int[] itemArr = findIndexTable(oriTable, matchTable);
                sameColorNum(sDataGridView, itemArr, oriTable, matchTable);
            }
            else
            {
                //numColumns test = (numColumns)num;
                //	進行連結資料庫
                conn.Open();

                //oledbdataadapter物件建立資料表查詢結果
                //	宣告並設定  查詢『num 資料表』字串
                string str = "select 日期,一,二,三,四,五,六,特 from num where "
                    + "(一 = " + sNumTextBox.Text + " or 二 = " + sNumTextBox.Text
                    + " or 三 = " + sNumTextBox.Text + " or 四 = " + sNumTextBox.Text
                    + " or 五 = " + sNumTextBox.Text + " or 六 = " + sNumTextBox.Text
                    + " or 特 = " + sNumTextBox.Text
                    + ") and 日期 between " + sDateTextBox1.Text
                    + " and " + sDateTextBox2.Text + " order by 日期";

                //	宣告並設定  資料表查詢物件『adapter』
                adapter = new OleDbDataAdapter(str, conn);

                DataTable matchTable = new DataTable();
                DataTable oriTable = new DataTable();
                //	將伺服器資料庫的查詢結果（adapter）存放到table上
                adapter.Fill(matchTable);

                adapter = new OleDbDataAdapter("select * from num where 日期 between " + sDateTextBox1.Text + " and " + sDateTextBox2.Text + " order by 日期", conn);
                adapter.Fill(oriTable);
                //關閉連線
                conn.Close();

                //filterTable(oriTable, matchTable, 0);
                sDataGridView.DataSource = matchTable;
                sDataGridView.ClearSelection();
                changeColor(sDataGridView, Convert.ToInt16(sNumTextBox.Text), Color.Yellow);

                //20100704 refined
                qrForm.totalTable.Merge(matchTable);
                qrForm.refresh();
                qrForm.Show();

                //int[] itemArr = findIndexTable(oriTable, matchTable);
                //sameColorNum(sDataGridView, itemArr, oriTable, matchTable);
            }
            //MessageBox.Show("不合法的輸入");
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
        //將傳入之兩row之內含值相同者塗上color
        private void paint(DataGridView dgv, DataRow row1, DataRow row2, Color color)
        {
            int m = -1, n = -1;
            for (int i = 0; i < dgv.Rows.Count; i++)
                if (dgv.Rows[i].Cells["日期"].Value.Equals(row1["日期"]))
                    m = i;
                else if (dgv.Rows[i].Cells["日期"].Value.Equals(row2["日期"]))
                    n = i;

            for (int i = 1; i <= 7; i++)
                for (int j = 1; j <= 7; j++)
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
            for (int i = 0; i <= 7; i++)
                dgv.Rows[m].Cells[i].Style.BackColor = color;
        }
        //sort
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
        //table1為原始資料,table2為部分集合,此方法將table1中符合table2下n期之資料加入table2中
        private void filterTable(DataTable table1, DataTable table2, int n)
        {
            int[] arr = findIndexTable(table1, table2);

            for (int i = 0; i < arr.Length; i++)
            {
                if (n == 1)
                {
                    if ((arr[i] + n) < table1.Rows.Count)
                    {
                        if (arr.Contains(arr[i] + n))
                            continue;
                        else
                        {
                            int temp = findRowIndex(table1.Rows[arr[i]], table2);

                            if (sameElement(table1.Rows[arr[i] + n], table2.Rows[temp]))
                                table2.ImportRow(table1.Rows[arr[i] + n]);
                            else if (arr.Contains(arr[i] - 1))
                                continue;
                            else
                            {
                                table2.Rows[temp].Delete();
                                table2.AcceptChanges();
                            }
                        }
                    }
                }
                else
                {
                    for (int j = 1; j <= n; j++)
                    {
                        if (((arr[i] + j) < table1.Rows.Count) && (!arr.Contains(arr[i] + j)))
                        {
                            if (!tableContainRow(table1.Rows[arr[i] + j], table2))
                                table2.ImportRow(table1.Rows[arr[i] + j]);
                            table2.AcceptChanges();
                        }
                    }
                }
            }
            table2.DefaultView.Sort = "日期";
            if (table2.Rows.Count == 1)
                table2.Rows[0].Delete();
            table2.AcceptChanges();
        }

        private bool tableContainRow(DataRow row, DataTable table)
        {
            for (int i = 0; i < table.Rows.Count; i++)
                if (table.Rows[i]["日期"].Equals(row["日期"]))
                    return true;
            return false;
        }
        //找出傳入row位於傳入table中的索引
        private int findRowIndex(DataRow row, DataTable table)
        {
            for (int i = 0; i < table.Rows.Count; i++)
                if (row["日期"].Equals(table.Rows[i]["日期"]))
                    return i;
            return -1;
        }
        //判斷傳入之row1和row2是否有相同的號碼
        private bool sameElement(DataRow row1, DataRow row2)
        {
            for (int i = 1; i <= 7; i++)
                for (int j = 1; j <= 7; j++)
                    if (row1[i].Equals(row2[j]))
                        return true;
            return false;
        }
        //找出table2中資料列位在table1中的索引
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

        //過濾出欲輸出之table
        private DataTable tableFilter(DataTable oriTable, DataTable mTable, int n, int preN)
        {
            DataTable oTable = oriTable.Copy();

            int[] arr = findIndexTable(oTable, mTable);
            bool[] filter = new bool[oTable.Rows.Count];

            for (int i = 0; i < oTable.Rows.Count; i++)
                for (int j = 0; j < arr.Length; j++)
                    if (((arr[j] + n) < oTable.Rows.Count) && ((arr[j] - preN) >= 0))
                        if ((i >= (arr[j] - preN)) && (i <= (arr[j] + n)))
                        {
                            filter[i] = true;
                            break;
                        }
            for (int i = 0; i < oTable.Rows.Count; i++)
                if (filter[i] == false)
                    oTable.Rows[i].Delete();
            oTable.AcceptChanges();
            if (oTable.Rows.Count == 2)
                return null;
            else
                return oTable;
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

        //舊查詢按鈕
        private void query2Button_Click(object sender, EventArgs e)
        {

            //	進行連結資料庫
            conn.Open();
            //oledbdataadapter物件建立資料表查詢結果
            //	宣告並設定  查詢『num 資料表』字串
            string str = "select * from num where 日期 between " + q2DateTextBox1.Text + " and " + q2DateTextBox2.Text + " order by 日期";

            //	宣告並設定  資料表查詢物件『adapter』
            adapter = new OleDbDataAdapter(str, conn);

            DataTable table = new DataTable();
            adapter.Fill(table);

            //關閉連線
            conn.Close();

            if (table.Rows.Count != 0)
            {

                //宣告並判斷後幾期變數,預設為0
                int rangeNum = 0;
                if (rangeTextBox2.Text == "")
                {
                    rangeNum = 0;
                    rangeTextBox2.Text = rangeNum.ToString();
                }
                else
                    rangeNum = Int32.Parse(rangeTextBox2.Text);

                bool[] status = new bool[8];
                int[] inputNum = new int[7];

                //判斷是否輸入錯誤(未輸入所有查詢值)
                if ((textBox2_1.Text == "") && (textBox2_2.Text == "") && (textBox2_3.Text == "") && (textBox2_4.Text == "") && (textBox2_5.Text == "") && (textBox2_6.Text == "") && (textBox2_7.Text == ""))
                    status[0] = false;
                else
                    status[0] = true;

                //判斷輸入值
                //position 1
                if (textBox2_1.Text == "")
                    status[1] = true;
                else
                    inputNum[0] = Int32.Parse(textBox2_1.Text);
                //position 2
                if (textBox2_2.Text == "")
                    status[2] = true;
                else
                    inputNum[1] = Int32.Parse(textBox2_2.Text);
                //position 3
                if (textBox2_3.Text == "")
                    status[3] = true;
                else
                    inputNum[2] = Int32.Parse(textBox2_3.Text);
                //position 4
                if (textBox2_4.Text == "")
                    status[4] = true;
                else
                    inputNum[3] = Int32.Parse(textBox2_4.Text);
                //position 5
                if (textBox2_5.Text == "")
                    status[5] = true;
                else
                    inputNum[4] = Int32.Parse(textBox2_5.Text);
                //position 6
                if (textBox2_6.Text == "")
                    status[6] = true;
                else
                    inputNum[5] = Int32.Parse(textBox2_6.Text);
                //position 7
                if (textBox2_7.Text == "")
                    status[7] = true;
                else
                    inputNum[6] = Int32.Parse(textBox2_7.Text);

                if (status[0])
                {

                    for (int i = table.Rows.Count - 1; i - rangeNum >= 0; i--)
                    {
                        if ((status[1] || (System.Convert.ToInt32(table.Rows[i - rangeNum][1]) == inputNum[0])) && (status[2] || (System.Convert.ToInt32(table.Rows[i - rangeNum][2]) == inputNum[1])) && (status[3] || (System.Convert.ToInt32(table.Rows[i - rangeNum][3]) == inputNum[2])) && (status[4] || (System.Convert.ToInt32(table.Rows[i - rangeNum][4]) == inputNum[3])) && (status[5] || (System.Convert.ToInt32(table.Rows[i - rangeNum][5]) == inputNum[4])) && (status[6] || (System.Convert.ToInt32(table.Rows[i - rangeNum][6]) == inputNum[5])) && (status[7] || (System.Convert.ToInt32(table.Rows[i - rangeNum][7]) == inputNum[6])))
                            continue;
                        else
                            table.Rows[i].Delete();
                    }
                    for (int i = 0; i < rangeNum; i++)
                        table.Rows[i].Delete();

                    table.AcceptChanges();

                    objDV2.DataSource = table;

                    //20100704 refined
                    qrForm.totalTable.Merge(table);
                    qrForm.refresh();
                    qrForm.Show();


                    objDV2.ClearSelection();
                    if (table.Rows.Count != 0)
                    {
                        for (int i = 0; i < table.Rows.Count; i++)
                            for (int j = 1; j <= 7; j++)
                                totalNum[Convert.ToInt32(table.Rows[i][j]) - 1]++;
                    }
                }
                else
                    MessageBox.Show("請輸入欲查詢數字！");
            }
            else
                MessageBox.Show("查無此日期範圍資料...");
        }

        private void showTotalResult(int[] num)
        {
            int[] max = new int[10];
            max = findTopNIndex(num, 10);
            string outStr = "號碼\t次數";
            for (int i = 0; i < 10; i++)
                if ((max[i] + 1) < 10)
                {
                    if (num[max[i]] < 10)
                        outStr += '\n' + "   0" + (max[i] + 1).ToString() + "\t" + "   0" + num[max[i]].ToString();
                    else
                        outStr += '\n' + "   0" + (max[i] + 1).ToString() + '\t' + "   " + num[max[i]].ToString();
                }
                else
                {
                    if (num[max[i]] < 10)
                        outStr += '\n' + "   " + (max[i] + 1).ToString() + '\t' + "   0" + num[max[i]].ToString();
                    else
                        outStr += '\n' + "   " + (max[i] + 1).ToString() + '\t' + "   " + num[max[i]].ToString();
                }
            MessageBox.Show(outStr);
        }

        private void objDV2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1 && e.ColumnIndex != 0)
            {
                clearColor(objDV2);
                objDV2.CurrentCell.Style.BackColor = Color.Yellow;
                changeColor(objDV2, Convert.ToInt32(objDV2.CurrentCell.Value), Color.Yellow);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Array.Clear(totalNum, 0, totalNum.Length);
            showTotalResult(totalNum);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            showTotalResult(totalNum);
        }

        private void pButton_Click(object sender, EventArgs e)
        {
            //	進行連結資料庫
            conn.Open();
            //oledbdataadapter物件建立資料表查詢結果
            //	宣告並設定  查詢『num 資料表』字串
            string str = "select * from num where 日期 between " + pDateTextBox1.Text + " and " + pDateTextBox2.Text + " order by 日期";

            //	宣告並設定  資料表查詢物件『adapter』
            adapter = new OleDbDataAdapter(str, conn);

            DataTable table = new DataTable();
            adapter.Fill(table);

            //關閉連線
            conn.Close();

            if (table.Rows.Count != 0)
            {
                //宣告並判斷後幾期變數,預設為1
                int rangeNum = 1;
                if (pRangeTextBox.Text == "")
                {
                    rangeNum = 1;
                    pRangeTextBox.Text = rangeNum.ToString();
                }
                else
                    rangeNum = Int32.Parse(pRangeTextBox.Text);

                bool[] status = new bool[8];
                int[] inputNum = new int[7];

                //預測統計中之數字陣列
                int[] pTotalNum = new int[49];

                //判斷是否輸入錯誤(未輸入所有查詢值)
                if ((pTextBox1.Text == "") && (pTextBox2.Text == "") && (pTextBox3.Text == "") && (pTextBox4.Text == "") && (pTextBox5.Text == "") && (pTextBox6.Text == "") && (pTextBox7.Text == ""))
                    status[0] = false;
                else
                    status[0] = true;

                //判斷輸入值
                //position 1
                if (pTextBox1.Text == "")
                    status[1] = true;
                else
                    inputNum[0] = Int32.Parse(pTextBox1.Text);
                //position 2
                if (pTextBox2.Text == "")
                    status[2] = true;
                else
                    inputNum[1] = Int32.Parse(pTextBox2.Text);
                //position 3
                if (pTextBox3.Text == "")
                    status[3] = true;
                else
                    inputNum[2] = Int32.Parse(pTextBox3.Text);
                //position 4
                if (pTextBox4.Text == "")
                    status[4] = true;
                else
                    inputNum[3] = Int32.Parse(pTextBox4.Text);
                //position 5
                if (pTextBox5.Text == "")
                    status[5] = true;
                else
                    inputNum[4] = Int32.Parse(pTextBox5.Text);
                //position 6
                if (pTextBox6.Text == "")
                    status[6] = true;
                else
                    inputNum[5] = Int32.Parse(pTextBox6.Text);
                //position 7
                if (pTextBox7.Text == "")
                    status[7] = true;
                else
                    inputNum[6] = Int32.Parse(pTextBox7.Text);

                if (status[0])
                {
                    for (int i = table.Rows.Count - 1; i - rangeNum >= 0; i--)
                    {
                        //if ((status[1] || (System.Convert.ToInt32(table.Rows[i - rangeNum][1]) == inputNum[0])) && (status[2] || (System.Convert.ToInt32(table.Rows[i - rangeNum][2]) == inputNum[1])) && (status[3] || (System.Convert.ToInt32(table.Rows[i - rangeNum][3]) == inputNum[2])) && (status[4] || (System.Convert.ToInt32(table.Rows[i - rangeNum][4]) == inputNum[3])) && (status[5] || (System.Convert.ToInt32(table.Rows[i - rangeNum][5]) == inputNum[4])) && (status[6] || (System.Convert.ToInt32(table.Rows[i - rangeNum][6]) == inputNum[5])) && (status[7] || (System.Convert.ToInt32(table.Rows[i - rangeNum][7]) == inputNum[6])))
                        if (containNums(table.Rows[i - rangeNum], inputNum, false))
                            continue;
                        else
                            table.Rows[i].Delete();
                    }
                    for (int i = 0; i < rangeNum; i++)
                        table.Rows[i].Delete();

                    table.AcceptChanges();

                    pDataGridView.DataSource = table;
                    pDataGridView.ClearSelection();

                    //20100704 refined
                    qrForm.totalTable.Merge(table);
                    qrForm.refresh();
                    qrForm.Show();

                    if (table.Rows.Count != 0)
                    {
                        for (int i = 0; i < table.Rows.Count; i++)
                            for (int j = 1; j <= 7; j++)
                                pTotalNum[Convert.ToInt32(table.Rows[i][j]) - 1]++;
                    }

                    showTotalResult(pTotalNum);
                }
                else
                {
                    MessageBox.Show("請輸入欲查詢數字！");
                    pDataGridView.DataSource = null;
                }
            }
            else
                MessageBox.Show("查無此日期範圍資料...");
        }

        private void pDataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1 && e.ColumnIndex != 0)
            {
                clearColor(pDataGridView);
                pDataGridView.CurrentCell.Style.BackColor = Color.Yellow;
                changeColor(pDataGridView, Convert.ToInt32(pDataGridView.CurrentCell.Value), Color.Yellow);
            }
        }

        private void test()
        {
            //	進行連結資料庫
            conn.Open();
            //oledbdataadapter物件建立資料表查詢結果
            //	宣告並設定  查詢『num 資料表』字串
            string str = "select * from num  order by 日期";

            //	宣告並設定  資料表查詢物件『adapter』
            adapter = new OleDbDataAdapter(str, conn);

            DataTable oriTable = new DataTable();
            adapter.Fill(oriTable);

            //關閉連線
            conn.Close();

            DataTable outTable = new DataTable();

            for (int k = oriTable.Rows.Count - 2; k >= 0; k--)
            {
                DataRow row = oriTable.Rows[k];

                int[] inputNum = new int[7];
                int[] max = new int[6];
                int[] resultNum = new int[49];

                for (int j = 0; j < 7; j++)
                    inputNum[j] = Convert.ToInt32(row[j + 1]);

                DataTable table = oriTable.Copy();

                for (int i = table.Rows.Count - 1; i - 1 >= 0; i--)
                {
                    if (containNums(table.Rows[i - 1], inputNum, false))
                        continue;
                    else
                        table.Rows[i].Delete();
                }
                for (int i = 0; i < 1; i++)
                    table.Rows[i].Delete();

                table.AcceptChanges();

                if (table.Rows.Count != 0)
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                        for (int j = 1; j <= 7; j++)
                            resultNum[Convert.ToInt32(table.Rows[i][j]) - 1]++;
                }
                showTotalResult2(inputNum, resultNum);
            }

        }

        private void showTotalResult2(int[] input, int[] num)
        {
            int[] max = new int[7];
            max = findTopNIndex(num, 7);
            string outStr = "輸入號碼:";
            for (int i = 0; i < 7; i++)
                if (input[i] < 10)
                    outStr += "0" + input[i].ToString() + " ";
                else
                    outStr += input[i].ToString() + " ";
            outStr += "\n前七號碼:";
            for (int i = 0; i < 7; i++)
                if ((max[i] + 1) < 10)
                    outStr += "0" + (max[i] + 1).ToString() + " ";
                else
                    outStr += (max[i] + 1).ToString() + " ";

            MessageBox.Show(outStr);
        }

        private void aQueryButton_Click(object sender, EventArgs e)
        {
            DataTable table = new DataTable();
            DataTable result1 = new DataTable();
            DataTable result2 = new DataTable();

            //	進行連結資料庫
            conn.Open();

            //	宣告並設定  查詢『num 資料表』字串
            string str = "select * from num where 日期 between " + aQueryDateTextBox1.Text + " and " + aQueryDateTextBox2.Text + " order by 日期";
            adapter = new OleDbDataAdapter(str, conn);
            adapter.Fill(table);

            str = "select * from result1";
            adapter = new OleDbDataAdapter(str, conn);
            adapter.Fill(result1);

            str = "select * from result2";
            adapter = new OleDbDataAdapter(str, conn);
            adapter.Fill(result2);

            //關閉連線
            conn.Close();

            int[] inputNum = new int[2];

            if (numTextBox1.Text != "" && numTextBox2.Text != "")
            {
                inputNum[0] = Convert.ToInt16(numTextBox1.Text);
                inputNum[1] = Convert.ToInt16(numTextBox2.Text);
            }
            else
                MessageBox.Show("請輸入號碼...");

            //judge if the row contains inputNum
            int count = 0;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                int[] tempNum = new int[7];

                for (int j = 1; j <= 7; j++)
                {
                    tempNum[j - 1] = Convert.ToInt32(table.Rows[i][j]);
                }

                if (tempNum.Contains(inputNum[0]) && tempNum.Contains(inputNum[1]))
                {
                    count++;
                }

            }

            //save rest numbers in matching numbers
            int[] numArray = new int[count * 5];
            int n = 0;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                int[] tempNum = new int[7];

                for (int j = 1; j <= 7; j++)
                {
                    tempNum[j - 1] = Convert.ToInt32(table.Rows[i][j]);
                }

                if (tempNum.Contains(inputNum[0]) && tempNum.Contains(inputNum[1]))
                {
                    for (int k = 0; k < 7; k++)
                    {
                        if (tempNum[k] != inputNum[0] && tempNum[k] != inputNum[1])
                        {
                            numArray[n++] = tempNum[k];
                        }
                    }
                }

            }

            //記錄號碼與尾數出現次數
            int[] numCount = new int[49];
            int[] tail = new int[10];

            for (int i = 0; i < numArray.Length; i++)
            {
                int temp = numArray[i];
                numCount[temp - 1]++;
                tail[temp % 10]++;
            }

            int[] topFiveTail = findTopNIndex(tail, 5);
            int[] topTenNum = findTopNIndex(numCount, 10);

            //result1 process
            for (int i = 0; i < result1.Rows.Count; i++)
            {
                result1.Rows[i]["號碼"] = topTenNum[i] + 1;
                result1.Rows[i]["次數"] = numCount[topTenNum[i]];
                result1.Rows[i]["機率"] = numCount[topTenNum[i]] / 49.0 / 7.0;
            }


            //result2 process
            for (int i = 0; i < result2.Rows.Count; i++)
            {
                result2.Rows[i]["尾數"] = topFiveTail[i];
                int[] temp = new int[5];

                for (int j = 0; j < 5; j++)
                    if (topFiveTail[i] == 0)
                    {
                        if (j == 0)
                            temp[j] = -1;
                        else
                            temp[j] = numCount[topFiveTail[i] - 1 + 10 * j];
                    }
                    else
                        temp[j] = numCount[topFiveTail[i] - 1 + 10 * j];

                int[] top3 = findTopNIndex(temp, 3);
                result2.Rows[i]["號碼一"] = 10 * top3[0] + Convert.ToInt16(result2.Rows[i]["尾數"]);
                result2.Rows[i]["號碼二"] = 10 * top3[1] + Convert.ToInt16(result2.Rows[i]["尾數"]);
                result2.Rows[i]["號碼三"] = 10 * top3[2] + Convert.ToInt16(result2.Rows[i]["尾數"]);
            }

            aQueryDataGridView1.DataSource = result1;
            aQueryDataGridView2.DataSource = result2;
            aQueryDataGridView1.ClearSelection();
            aQueryDataGridView2.ClearSelection();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            DataTable table = new DataTable();
            DataTable resultTable = new DataTable();

            //	進行連結資料庫
            conn.Open();

            //	宣告並設定  查詢『num 資料表』字串
            string str = "select * from num order by 日期";//where 日期 between " + aQueryDateTextBox1.Text + " and " + aQueryDateTextBox2.Text + " order by 日期";
            adapter = new OleDbDataAdapter(str, conn);
            adapter.Fill(table);

            //關閉連線
            conn.Close();

            int pos1 = 0, pos2 = 0, n = Convert.ToInt32(nextTextBox.Text);

            if (radioButton1_1.Checked)
                pos1 = 1;
            else if (radioButton1_2.Checked)
                pos1 = 2;
            else if (radioButton1_3.Checked)
                pos1 = 3;
            else if (radioButton1_4.Checked)
                pos1 = 4;
            else if (radioButton1_5.Checked)
                pos1 = 5;
            else if (radioButton1_6.Checked)
                pos1 = 6;
            else if (radioButton1_7.Checked)
                pos1 = 7;

            if (radioButton2_1.Checked)
                pos2 = 1;
            else if (radioButton2_2.Checked)
                pos2 = 2;
            else if (radioButton2_3.Checked)
                pos2 = 3;
            else if (radioButton2_4.Checked)
                pos2 = 4;
            else if (radioButton2_5.Checked)
                pos2 = 5;
            else if (radioButton2_6.Checked)
                pos2 = 6;
            else if (radioButton2_7.Checked)
                pos2 = 7;

            resultTable = table.Copy();
            resultTable.Clear();

            for (int i = 0; i < table.Rows.Count - 2; i++)
            {
                if (Convert.ToInt32(table.Rows[i][pos1]) == Convert.ToInt32(table.Rows[i + 1][pos2]))
                {
                    for (int j = 0; j < 2; j++)
                        resultTable.ImportRow(table.Rows[i + j]);
                    if ((i + 2 + n) < table.Rows.Count)
                        resultTable.ImportRow(table.Rows[i + 2 + n]);
                }
            }

            pQueryDataGridView.DataSource = resultTable;

            //20100704 refined
            qrForm.totalTable.Merge(resultTable);
            qrForm.refresh();
            qrForm.Show();

            for (int k = 0; k < resultTable.Rows.Count; k = k + 3)
            {
                for (int i = 1; i <= 7; i++)
                    for (int j = 1; j <= 7; j++)
                        if (pQueryDataGridView.Rows[k].Cells[i].Value.Equals(pQueryDataGridView.Rows[k + 1].Cells[j].Value))
                        {
                            pQueryDataGridView.Rows[k].Cells[i].Style.BackColor = Color.Pink;
                            pQueryDataGridView.Rows[k + 1].Cells[j].Style.BackColor = Color.Pink;
                            break;
                        }
                paintBack(pQueryDataGridView, resultTable.Rows[k + 2], Color.Gray);
            }
        }

        private void dayQueryButton_Click(object sender, EventArgs e)
        {
            DataTable table = new DataTable();
            DataTable resultTable = new DataTable();

            //	進行連結資料庫
            conn.Open();

            //	宣告並設定  查詢『num 資料表』字串
            string str = "select * from num order by 日期";//where 日期 between " + aQueryDateTextBox1.Text + " and " + aQueryDateTextBox2.Text + " order by 日期";
            adapter = new OleDbDataAdapter(str, conn);
            adapter.Fill(table);

            //關閉連線
            conn.Close();

            resultTable = table.Copy();
            resultTable.Clear();

            for (int i = 0; i < table.Rows.Count; i++)
            {
                if ((Convert.ToInt32(table.Rows[i][0]) % 100) == Convert.ToInt32(dayTextBox.Text))
                    resultTable.ImportRow(table.Rows[i]);
            }

            dateDataGridView.DataSource = resultTable;

            //20100704 refined
            qrForm.totalTable.Merge(resultTable);
            qrForm.refresh();
            qrForm.Show();
        }

        private void monthQueryButton_Click(object sender, EventArgs e)
        {
            DataTable table = new DataTable();
            DataTable resultTable = new DataTable();

            //	進行連結資料庫
            conn.Open();

            //	宣告並設定  查詢『num 資料表』字串
            string str = "select * from num order by 日期";//where 日期 between " + aQueryDateTextBox1.Text + " and " + aQueryDateTextBox2.Text + " order by 日期";
            adapter = new OleDbDataAdapter(str, conn);
            adapter.Fill(table);

            //關閉連線
            conn.Close();

            resultTable = table.Copy();
            resultTable.Clear();

            int j = Convert.ToInt32(monthTextBox.Text);
            int temp = 0, index = -1, counter = 0;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (temp != (Convert.ToInt32(table.Rows[i][0]) / 100))
                {
                    temp = Convert.ToInt32(table.Rows[i][0]) / 100;
                    if (counter >= j)
                        resultTable.ImportRow(table.Rows[index + j - 1]);
                    counter = 0;
                    index = i;
                    counter++;
                }
                else
                {
                    counter++;
                }
            }

            dateDataGridView.DataSource = resultTable;

            //20100704 refined
            qrForm.totalTable.Merge(resultTable);
            qrForm.refresh();
            qrForm.Show();
        }

        private void dateDataGridView_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex != -1 && e.ColumnIndex != 0)
            {
                clearColor(dateDataGridView);
                dateDataGridView.CurrentCell.Style.BackColor = Color.Yellow;
                changeColor(dateDataGridView, Convert.ToInt32(dateDataGridView.CurrentCell.Value), Color.Yellow);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DataTable table = new DataTable();
            DataTable resultTable = new DataTable();

            //	進行連結資料庫
            conn.Open();

            //	宣告並設定  查詢『num 資料表』字串
            string str = "select * from num order by 日期";//where 日期 between " + aQueryDateTextBox1.Text + " and " + aQueryDateTextBox2.Text + " order by 日期";
            adapter = new OleDbDataAdapter(str, conn);
            adapter.Fill(table);

            //關閉連線
            conn.Close();

            resultTable = table.Copy();
            resultTable.Clear();

            int temp = Convert.ToInt32(table.Rows[0][0]) / 100;
            for (int i = 1; i < table.Rows.Count; i++)
            {
                if (temp != (Convert.ToInt32(table.Rows[i][0]) / 100))
                {
                    resultTable.ImportRow(table.Rows[i - 1]);
                    temp = Convert.ToInt32(table.Rows[i][0]) / 100;
                }
            }

            dateDataGridView.DataSource = resultTable;

            //20100704 refined
            qrForm.totalTable.Merge(resultTable);
            qrForm.refresh();
            qrForm.Show();
        }
    }
}
