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
    public partial class QueryResultForm : Form
    {
        string strConn = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=LT.mdb";
        OleDbConnection conn;
        OleDbDataAdapter adapter;
        public DataTable result1;
        public DataTable result2;
        public DataTable totalTable;

        public QueryResultForm()
        {
            InitializeComponent();
            function();
            dataGridView1.ClearSelection();
            dataGridView2.ClearSelection();
        }

        private void function()
        {
            conn=new OleDbConnection(strConn);
            result1 = new DataTable();
            result2 = new DataTable();
            totalTable = new DataTable();

            String str = "select * from result1";
            adapter = new OleDbDataAdapter(str, conn);
            adapter.Fill(result1);

            str = "select * from result2";
            adapter = new OleDbDataAdapter(str, conn);
            adapter.Fill(result2);

            dataGridView1.DataSource = result1;
            dataGridView2.DataSource = result2;
        }

        private void clearButton_Click(object sender, EventArgs e)
        {
            totalTable.Clear();
            refresh();
        }

        private void deleteSameRow()
        {
            
        }

        public void refresh()
        {
            deleteSameRow();
            //記錄號碼與尾數出現次數
            int[] count = new int[49];
            int[] tail = new int[10];

            for (int i = 0; i < totalTable.Rows.Count; i++)
                for (int j = 1; j < totalTable.Columns.Count; j++)
                {
                    int temp = System.Convert.ToInt32(totalTable.Rows[i][j]);
                    count[temp - 1]++;
                    tail[temp % 10]++;
                }

            int[] topFiveTail = findTopNIndex(tail, 5);
            int[] topTenNum = findTopNIndex(count, 10);


            //result1 process
            for (int i = 0; i < result1.Rows.Count; i++)
            {
                result1.Rows[i]["號碼"] = topTenNum[i] + 1;
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
        }

        public int[] findTopNIndex(int[] num, int n)
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

        private void QueryResultForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
        }
    }
}
