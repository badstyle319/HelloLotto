using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace DailyCash
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
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            conn.Open();
            string sql = "create table dailycash (日期 int, " +
            "一 int, 二 int, 三 int, 四 int, 五 int)";
            OleDbCommand cmd = new OleDbCommand(sql, conn);
            cmd.ExecuteNonQuery();
            conn.Close();

            refreshData();
            gotoRow(currentRow);
            updateDate();
        }

        //更新資料庫內容到異動datagrid上
        private void refreshData()
        {
            datasetNum.Clear();
            //	進行連結資料庫
            conn.Open();
            //oledbdataadapter物件建立資料表查詢結果
            //	宣告並設定  查詢『num 資料表』字串
            string str = "select * from dailycash order by 日期";
            //	宣告並設定  資料表查詢物件『adapter』
            adapter = new OleDbDataAdapter(str, conn);

            //	將伺服器資料庫的查詢結果（adapter）存放並填滿到終端機的暫存物件（datasetNum）上的表格"num"
            adapter.Fill(datasetNum, "dailycash");

            //在datagrid來顯示dataset上的資料
            objDV.DataSource = datasetNum.Tables["dailycash"];
            int nRowNum = datasetNum.Tables["dailycash"].Rows.Count;
            if (nRowNum != 0)
                newestDate = datasetNum.Tables["dailycash"].Rows[nRowNum - 1]["日期"].ToString();
            //關閉連線
            conn.Close();

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
            okButton.Enabled = false;
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            firstButton.Enabled = true;
            preButton.Enabled = true;
            nextButton.Enabled = true;
            lastButton.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            enableEdit();
            dateTextBox.Text = "";
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            editStatus = 1;
            dateTextBox.Focus();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            enableEdit();
            editStatus = 2;
            button1.Enabled = false;
            button3.Enabled = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult answer = MessageBox.Show("您確定要刪除此筆資料嗎？", "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);

            if (answer == DialogResult.Yes)
            {
                conn.Open();

                //刪除資料庫內的記錄
                //	設定刪除記錄的 SQL語法 及資料庫執行指令OleDbCommand
                string str = "delete * from dailycash where 日期=" + Int32.Parse(dateTextBox.Text) + "";
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
                    string str = "Insert Into dailycash(日期,一,二,三,四,五)Values(" + Int32.Parse(dateTextBox.Text) + "," + Int32.Parse(textBox1.Text) + "," + Int32.Parse(textBox2.Text) + "," + Int32.Parse(textBox3.Text) + "," + Int32.Parse(textBox4.Text) + "," + Int32.Parse(textBox5.Text) + ")";

                    OleDbCommand cmd = new OleDbCommand(str, conn);

                    //執行資料庫指令OleDbCommand
                    cmd.ExecuteNonQuery();

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
                    string str = "Update dailycash set 日期 = " + Int32.Parse(dateTextBox.Text) + ",一=" + Int32.Parse(textBox1.Text) + ",二=" + Int32.Parse(textBox2.Text) + ",三=" + Int32.Parse(textBox3.Text) + ",四=" + Int32.Parse(textBox4.Text) + ",五=" + Int32.Parse(textBox5.Text) + " where 日期= " + Int32.Parse(dateTextBox.Text) + "";

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
    }
}
