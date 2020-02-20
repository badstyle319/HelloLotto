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
            for (int nMonth = 1; nMonth <= 1; nMonth++)
            {
                nvc.Clear();
                //nvc.Add("__VIEWSTATE", "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");
                nvc.Add("__VIEWSTATE", "WvCBw068nOMANxqjC2TrcE+PGEqORElQF60VH2qdui/a6wEGE+3TK7h1AewArszq8e6SBg8Dbwe93VM06SEQY1ZKx6qNqwChgwbHOf+WyHQ8Lt9jjHXszFo/yO7ppgnipcXpykJ5FD0hgk9Y4gsRoyp5DVqMk8jmvMM20xOXdwO6jrBXPcJJHtQw61tKOspRcxQlFSmslIYpGv+QZ2BcYEucmu/qqDnUxIjRmdP6VLtLtUoP7D899f9vyyiuFUygWPlju4paEfS8/j/IxVs/MnXka8FU+qLoIlECNuU7rgtGZyaXIcvV2Czmgc4QiC9OA7+GLy+K0oOGgrYziQeqmYKUl/wNSJGoBwGAe3BpA8tHqns/BGZ7eTyAzFHnkIXGK6WPWhj7UOGnU05hhND3aXipQ6v9QhQCO87NqmNygDnLy92US5viRQNdhND/HWhqRc+ewF7fNszr3H8RBjpXwQyOe72BcU0cyv+EPcdeNMbF1MAoTrJQwi1baLzFv/63lw3gRUDgvQ1HcZo6a/VS2eNVwyM1PldxC/8YDsrKqq1anAIVs2X7QjosrWZ4/HmHwUk/6cxw88Zb2bVKHeGJkEaPi2fQ2fa9um2oCfE4p97SVs0coChIL+8saInVnuFNkdkTVhbppbDESpKmS/b2l7RvrXkSpk58AYew2djUJkNQT4N8Ht5k3qVr9XY2ePzuhluJehXvM6Ahfi7jsfEIHQla+4NxPgA1Ev1yBcppY/SY9eLW1irYlQ53ry3KQmQMvXLATr+yYacOjrnDqKjxoaHJhs0bBsCwuee/4GwxXiUrAfmCEyt2bs8OganIroiS8uaYvLFN5mQBb88AaLjPl1QD6EUSmmqkrqcOfdaOMlvC7IZI5NZz4tZSZlWLjx6Uxz9sO7qO8bbuvlQI8HoopXZGhfVP+HhNxAHfJ9oKQ9Uhq/4v0BISsleY/devFO+mIM21E+57RvT8ATU0vx+Iny6hzK1Lx5wD8BhV2YQf9SYOiCdOLqGJ5yt6gTyN9Bfxmq1WluWGruibxcfgG1JQ2czUX/5zTlcJpu3cZwsXNnVrJDKOEpnPRc/6+1GW/gRPlADPsMAF8FcvSA1GYgRjcMMMtGO2xT1gzjjqTEHVwT39QBvWRMsBbZg7C3tmOJOozmMyfFRXxfQ15n8U6Kq+qozV7+/Vuo8tU25/xxR2wkf5N+kUCKrFFK+NgG3RCVXzEzd2fZEaNWys+QxUcY8LiKiO4OBV1gqrXlQ2r+5SuM9/XZY2BvVKMeUPeY5k7Kvk79cHv9jTesvZHd6CZXI3ovJQM2KCZ7tzfRc/AqN+6ublxk4gIKttJPC+v0FpMtqkhoFEXhEmHrV7fO7DOyF6mS1OmD6Z5FAM8PuYgZhiIKJj0Ak4t4+ciCqmEDOrt4DJK4mPX7wK0yOCLFYJuZJhH57ifRNcd4/H8j8Fo6WML1v05PfdHQ5Dn91Cr25/P16oPzqB4SgmbnHHHD9pL+P8jv5h3GXOHzcjGVUkfbUnRKcSbIEoWA7VGFlrKrfSUOFxZMEm/TDYI4+REWjnjUlF0RsMQGpyM1xzT4jg9csXrE7GCjh/7C6O6S8TfaYhOUOpSPFm1hyMatKrCD+He4yuh11mURmejVvVA4W5SRD5Fo8w8lcTbgfHD8jti5EBETW2RwbuvDk2VcBtgkGttMhi55nVUhTpC+csFUn4qg/+KqSVwCY4pm1hXW3ZK/e+ySyLzNmFDNtE4FkyY5mC0mJD8rpCdsDRowoGyONzRvnJpR/uuNvrXK4pP7447T3vTDkI2RLJDLA5VJO4F/I53QtNmIRegnhXsy5Nuy0BE0HKj9kWgtnmXh7q57lc5c6YOnJCoFNANlkv6c/IR4GDEneFhuFXMR4j3SzsaEz1hf6E4lCuD2YMJgWhVxPxmrgbDdC/+lPl9GJ4/zPwIouLs0kfjf4rACAaWrVu0uYgkv5qWeUxZZxIQ6yaQ/yL3ws6EXfzM3PYb93W+nN8L8HyecibXTjY9EvLnkvTTaN+7YeZRv08jh98nc26kelO6wQwfAZfmro7YsKb3BaGsjquWB6mJ43c8Wc+kfzpC6ggIMMstMu9DpYW9eMiOqrIDDqobW0me7Rsn5RRmSljs8VNAxH8L5FjjXgJC1TYtF8MOg51CVtqTZM/3A/LiBG3pvAN+J7c39JGrBqNfm+NRVT4pzmaoxNlDchb4xFsMFTsTu6edKJhsTivr8VeTvAYoTZk5YYQfBzgH/zs2uxgIRwXuhJI15XS9CpgN3bB7I7yITBbbJ6JpD1TLAY8p9OoRrmnhE6gHhP71/du9iODC4KiwdDeLimCCYbG4NCrc+UhLqRSPyxaOLbGSfOFBN9uFitYtdp1IvhmwAqg3qpb3/WYm5uWufGT42UKkE6GbS+IsZqV0iXvPm80f9UJOtiGQTbQJiflGJ1YCPEUFK5GpfUCPqtHjUdIRBif8pJYGTPzHHrr89DxvJNrRMHN6DyaMKaoMVzW8ZJl9EphNYAjPfv6BjcRdhBbaRN+h9aAWeVM+Uk2fBOq5rJDjbskMIWBVQ3H39mtyhy80DWg/iEE/Nmi2UXYe6UmHVUITKjin2F70PFbEeaVdF/Jrm5yQK5SZjaK/ltFWhsp8B2JKoE11tOe0ObWBYAtIdJV6n5lRia3rJeaLrtJz2RQIebwI+Ahq0/Olk5NLlgNrL3lHPJtj6wyowFWFkFFoGPI2JAVxXteh7mhvrBcDkdUtLHvw+TyTWGn+cDtgIHV0kITAx90h8nit1roeb0Na5St0xsiZ3it2ltxVZgMMKq3mHQ6yTBsj1poev9zGbAJrHfkylTzVctocG/wq4LOqcnZesgoF4jDe+Lbd0el7CuMKY6QFprEWk7XDXB42MzPX+GuT0x/I4pLHRRxFXTsq/kMAFgl/WQ2erZC2sOgy1ecFZXzobu05JICwsoTfkSnqWVfHh+xSO5rlYf17kslevVW6IlY3SzzmUP7SOIoJFVKT0l+zeFeE31oOTBKEW/5bAOFmJN4OV3qNCjaJCSJMTdSa4kmvK8CHaWimXtKOPzZGdIAGsJHXUWkXvh/DrHnuCPs60deQ934jtdKucNmutufqcKF5F7mTjDh/hVlSAcM+7C8Uv8V/U04YR6urKT4Szi3PvqLWrqiOCXWsByqc6dEmz3gJb0x0vfJ1OdaXRVaroE5L8e5xBAmVqtn4b6Xl6UaD2chX9ZNZOVFBLHZhZ0v5WOnlmE4G8zf12eEFP/4JAhZ7Yzfdad4CpkZNtdiBO7Pxbm/b44IVxeUuG37eqd1DRb9VL4QaObmYWcV0xHCvKMXwIWwTix95I34LAmAuG1aNidZZBBDUDvPmhaaTATRkahKnXJSauds2DYysaTXt9rHbGwIhi1k67ytSRx30+erLtvkX95gM6xS/VarKwA6uzJJL3y6lijoVxvkq0wbroPDOfIiNECsOhT89I/Onksrb7iUOv/AYXA66PRtpAHfCnNlaaUvduPENb7DhjpL4c43WtmknkU9N0QzPK96qC2X5f0FobBCjlrZa5kv/pBz0MwyR9yYHo703id1DdEtSDoTbY2bMs3Et6hIh9az3Y5ET9AN2RhM6FcJ9z/FIrbrCDL/TVNPMsFmejH5LQ61A/BD+Kh4N75Va7Hhx54lee5Atqco8427cMGpWb1cNggF9oVxYFyfhHMM19dHz4MuAxJDN18FBuAZqPsagCjmlDuEFOqNT1r7Dl73BXKS7WkQpULcnv2GaeKEtyNAetLcf61ZDzz73XlU6pJ4JVMwiyPhx0c+Rf6dV6iFhxuJLkkrqhSdPczrNoUr/AaFqK80N8Y/vt3GWx9oEhoyUyz73wy+MCVRlYNg14h6ptFLlakSpsE4f23Zt/q9FxIIJ5qhRAqNUq/BMWZqKNyu4k2hMDQPT6W16L0pZAjWOOUe1mp42F2Fac//so9VXdod07EdFN1MOSdNVikNLh9UEJxEgWTMW2Z8TYV2mlYHm5M/fv5SJ+eyz56YtY38LdF6Bp7WQFe1vL8H5gbeLVy01jCfpk+oF9Xksi8kODfTyBUZmOqJ3dCx9dvnQbSUBO8leF/am2uOqvCocVqZ2o/0WFDXd5PCHaJrk8vhUBk83rtjAxGV5wdcfO8nT94kA20Z67aTvx/XGbz7H7PWtW3MwBNLIvpoMDYNuQEZxErbK470eIuSWTKAIl5W7C0zgCrqcDLL6EXSbtbhPUm8NzC9zFPUXVsw4VHXVpYc01o0/tZ0xHFvvUHOlp3OBQ+G4VjEi0FZ44K+ZRD7v3jAOo7y3cCvAOKCTVf3bVmcb9FQSWmDgqCppn4b314xpoy+TrxeKS7r+1xwGIsWvdUwW1KQZWvid05VUVZd9cjARFw2v+K7huZoNAjzQdzvQQKJVARQnVLryttEbdaJuVChImE36PquEyXCzgGdgwOvCgME9gByKE5Sna1K+iVT4QYQoV95E/um2C3RxvtkimlrA5kmoDeljJHhNmgfEstYGhmR9Bd2OojUtRHdp3yFxbD2NmoFDAIXxD78LgVS+39tKpypcjE/646pF+/b1e5D/3/6ir32mwTP7aTsSDDLuDgI+JuuDhyErf5z/PWcMty2N0cAEt14cQlJGCPG0Se6XxISa1IIUlffcG/HyKAX47Hk7Ta0MNmtlddG2fDVsVZcsWMuPDVkNEmY4++H58Qn8cAO6URPco+qbzWck9CbzBI4mYadZvCSrbpi8UIMbX4i6+vDab0nShCF/EGco6ywnedO8rtFom1S0mbg6eQu+DtzNtNsVOtXKxavVVuiQ47Wvh/MnNcUTa+6qjiZcrD4zNJppiOQK+6A8mrRl/AUPGroyQSCR2hRpTrGZlfJMXXoeA5p4db8USEl/CfDL6W7YBSlOXDDI+ApOhyduZxcQmKjZJ0kbGw4LhgMvqIa4X3IZgfO7fnGPlscavX6mWlRsjvHMeyaZuiuRqO/h+bk2jFBEUO6NcLmeao5G3l8p76/JIVlcXj4IzdSbxzj2qogXdA12O3Bohy9dqhDuyLtZo5rf1Bjg+jtMBwcK9+pKkUHsxTfbv4wZLEepZ3Jy3IKTtB/Ty9ujYqkeJqcCZwXXDTAqflph6tHgvdJBp0ypMqiw7L3Wp9j2m5jWvIBjq8CSqxfvgx4pY54TTncbDSqW+W/COdgHW0tmIWUuMsnL44gqmZaXc9Da5Hu0Wj0Kvr5yzmlhmbJCnl7PTg1TK9VbCAo73YBJDshJA4jxaRYtGpuySmgdQWgY8K9qFE+uxjokfujARmA8XUc75evEFaApY7dT6s6qwR6QfGgUxuhNAy155CbFCWOVBaKhiMTW8EuHHJ9vNZJlexj/xsJS8LXjNaJBDGOmaRv5MFSQsuYscDCwvBH06CEzmIUFrL+7ZpLnE/9/wkVGCSu5YjtjOtg3IYBIr9Jws+HSDGzq6KYwyRJD1HYMNhYHe0ZyR9K8NUDSyqKSZFcidf8iDiuXDGzjlhJxGhwSgr53kAh4nb+BuOVaEA6Yr/VZN22lr4GhNDbbjE/VYEkvaPyfwpJGTdU8Hrjb10CNKWmq/4/95bn/WLkeFI96INyOAAxzQ+8HM3TnCYwWX4bKvluITWVu/39CUa4Rawee5TdAw0rLb5Ya35Kes/AOU68QantmTytrpfKv0AOSU01PYv20/XrO21QOWOvpmRzZdM8mqapUsBqlTBJbhdiPJGA744O4cYLXlXV3GkhCpyiTgE8e6rC8A+vxemFqx0pw0rRbCE7BAC2dms+eJdYESgJ0JjYbIE0+BeberA6BlttiTYTmETtInchj8oEL+MU5fhbx3I5jah/YXl/w2RmW9SJy5tFRLDFzUS/ePQdpkvnuaY8VMBNkf4K07oNovpvCgnCIq7iHmqw9148/In3rttadN16YYW6iNQUew3LdC0aES4Mu7ONjtplKMTDgt/60k+JF2/dr2nRW2Mf2J5vh6Jze4gPzJl6G30/DGYBoxx23Q8/Bs8+OdMWzvGkxjEWSsN5hu/5Ty220f+RTChPuodzipnvEyqi8DFJIeMM3zcoWuFJ21TnqXVIU5I2brx9j/2DyFGIt0IphzE6YkSCDcGVPYdxeuxfvcTP3ag5At96EFm/dp9Vu1AqCTMt+vvl89hWMPXj9j3/WeQQVlPkf5mW4rAXPLBxKeiW5+ITwj7MdeZhHZGxGCAxPHgGJ4iKSTT824WoXkPStqOA8sv9WvGDH2hkIA1E1DSVR8XsJK6gfPXy8rQ5oIpx554uRAhI6WiOI5GsSFPmomCSA0yovmnA3UTloW1thdt8qUo6cHmOr3oYM4t0zdF2P8MG5r6bm/oE6LjvYPsZw8lDwnmePUs5G5nDA/tiPI3/NfmGqMVwqJSxfpC6sxfkviOUWEsZw/qOzEPegZGqoIuUXTCAcnUQqJo1UWTtZ02Xa/OTnaD8rKytnSqsQtSFyHmZ7OvOEdduieppKuBi2zVEDzJz+kBinOoe7MJd8IP/8Qb4kMyRno5caI15N/+Cjyus+uupYBR6uBhK6K7P87IlHYR8xahmlBA7T8WBBZ7NF+yamEzcqN+38bw3UUXVAmmhkTOr1IJspc+9W7zbHXOZHPVMU9I0SqyKufYV5LC1wINt3RTTvf8eHk/uzNr4Geq3qYOaPxrnsyFSRBZv8FPXhxAzrENxqNYKaBq5r+i6xdafGoNeWPlHR5ZbyegYM3HnOTnPSx/+HCLV6wxiO5cy3GT6k1mXgq5Alc2KmVE6+u7cfHYU/pxGUS7H+B4ZDh3mdRUp7CWDu90X2RI/vstdFPo5NzFzuwBSnxwJqHj5I/rGq0fvkcozl2UV78sZArmQUicKgGVUJBnDoTvhVrnt0xVKboYkEpTQIlLpx02IeWR3UkF78Dm7CBO9wu10OO6W4VyoB9JJkWJI4rbZ70x53GuxD03jB60qb9JAYvVLCH7a06Ia0h7HoFNvKdcm2IZNPsoFb8evAEu730UYibzBSKVYHtDflTwIFmgrol9Sxpbf2wjbygAr9mQFGI3ckXae5IXOAHUYYH21UgCD6nwvA3KucW7PXlaGclVPWhAdOaTvRGYMUJGyAeIp5Dvnrh9OoG2/V/JrB//YXGVegl17wZtClTBBYYz+U/shAFjWAWVw5TtpKEA0qRKUnwR7/4nEKvwpXo3TwpgkTqSRtzSALobpfwznrKRnVX20mf9FP5lvaU92ZM1CoFRCSecqbN/XKPtb80TEkt4V1m+m0vAOpDOkZbt7QZAe6itfYRbHQB4fUTv2gOZj3Z9UuHY2v32O1Is9W6GKTnp8ygV9wRQoz53ATlLbLhI+acaDUNlYqqZkkSXjIkQsMkSZiVk5XRr4SRIFFSAqwg1/p1fEULx7Puj2QwcC9vaXUEz/PTkL16nkR03MmyCOcTCUsBGUgnUdoJPfoeXIxvsz9nYOZ/PKXq9zqmoxe5UaU+7KvaeqbmsmMilJioTwjhMeLtQ9Fimkfm5GkomnaWff1nUXTzVwugsRig51dSjyjeSdKu4psgXICEqnTF950du91FoIkaYBtrJExp8M3zPe9uFmx7a56r7Z4Ii9EGU1iLZNLZbaVziEH85hgiL1g82zM8IrFGXHNa+rYlaweaTK1G+0t0JtJx1Qn94q1KbeWfL5uvMhOIFjqN748xv1PyGMN9GMIdW6mA1Fpt17Llnp35lfNblCvpKGTCAJuKrk/xV00fMNNf2acqrVQMsnH5aPR1XCgx49ayKNOSlZcwg45JtCE5uaXF/K/Xmn9D2jXEVqunJG+0NSgsVE3JmQRjNemRNdVSwuco4MPx7xjlPD2IcmmLJ3ae7nDUSDTtGzWM9DeTrBzhGsa5kxQJJq5FTjpCdt/PK07alAzAbsgQ7RtVJs6ec8FMS9xJiVpu1CX7mpmK0N7zlUtv3xnM565tBoyEU5GxI6p0qxozztXCO2hzKPAIhtn7XI2VPttHJ6LxN8bI8GemMlcbSqnsZLw01z4RGQnFPbSGkr/53uwOTkMonYsO7IoHmjhr/dHLWZmFdSu0GFE3uHizQg1xnsKCK7NcohbaS4KA39Vbc42TP+QKDKnacvLc1evmKK+OSWwg2fKyd7FIwCIV4ub+fiGkCT3aJ3s8/btj+wLxMiWij28FczrIq/uFBUL52cBZczOovO55Yu/h8gWLRgDck0XKjIoPsl2oh9VxXH0w1lBX5prM4hxOodINwNRYsUp93677XHQWi9Z6y2HB9eu2OLNu8L6wmQlFAE61FHnVbo98oFTyxldYutytd2BoF3fnn9zgQnRkBARDkUjVrOAK2lpR5dpmv+RB46tKnS52G5rn2oIEW6Kkbq4vwEizMb/W28gWI87PhFPtpmd5szxMMz0oEymd8ren0iRo+iRdMMMGgBmGMPA5oX3luhWKX0gRrgLtp3uei4NIl482QQ4uO4+AuTxpSlL9T1OFS6m8XAnBe0uFeaiR954GmLkk2hmTdXJXhy18m3AmTJpUwCbchHeG/VdOJw/6sB1HCQL5JhFisWBy9X9z9ipv4TxL1zKbz//0PNXrjAvU6JbGYZjP+D68aQME2I29WBPFg8LrWm9ALU/YLwHZtoYpr5wYoRnQ1MKQ8MQ5fUwZ3Ph6yccttQHRaNx9mnDklWtvMpmFr1S9cMlyP04cKaNjM9cGGR0D4K8OkTNNTIrOG3ohwBnoR5Sa5Ar7u5Zc6rcCC7okrMNhdzlPS+Kdb5csrGb1Clqg==");
                nvc.Add("__VIEWSTATEGENERATOR", "09BD3138");
                //nvc.Add("__EVENTVALIDATION", "jLYlBZoPIVsYNIOrlZSYNj21iP5PGItrGiUIpXocb/cmsJvaddS6o/hXvc1KbV4AyVuqBwUjLBMaxuRuFzlhQ0kAF5p8tRN+YmaMzviQONDkweizlbbl86styjS9BonORP5tSpCc8urEvex5W0e6KD8oXNVX55mCjzkU6fpekl1MmL4Zv3g6bt+SSrmdASPKz/s3JLK510kK68rTBRGudUcS/6voVphIk35ZKHNPiq+VrZOvXUhSirBddiUg9jNMp5fDGYJOrxEYL5L14YtcnWWldlR5rCceK5G0/vRQ6ju9VCCXEkPizWdgcS18exNtSt4j1yUKE4HNP2YWqzfQ2kevLeoGiFHL5msDVhEvBWTPyrW3dVOK7M2sMcAgNTwEB5ERjrX49Xa5FYaRIAKM+Dz2JkzM9g8ZsilKX/GFgIASYhBDd0/i5cpvcPDfy4+6SKC+EY/Zrzzmb2VWaZ7b/kcLPUUGQyjVwbAl/qDcXsl7fPAT5MlmTWyxSoLxFcFODL1R7ssftHXgbYlXXYpJRaGuZq0VyquOvfAzPI0D+wUV2UrC2ZfdccYuWaWw8iRLUbmN4farVQPRV5HVPqosCBkVdEHHndiA/ETSjrpVZnHV0ETyDtOBx/NaMan5E8ZEZukZqBxJ7pLhSl8RGrtleMODTzPoX0WYg8euHtsyB3AyP/upZ3x1+MaMcHeIPSNsSxH8guOYefY1kCrsViI2SLX6v2SR6qKreB3sFLBLVRAU5orcg9sRu6YOEBV7lRLyR2xxBIU5OVSwE2eG9AOkQ54QIPSlW6xY0Rr8ZQXKy6ZCw0JFb64zur9L3Rmf5x4/uwn+BcMaV49h5TUrMAql6AbKmtI=");
                nvc.Add("__EVENTVALIDATION", "RXQxSgm9P48+nCZJy081ZfldviZiSdlgz84ME7sMrZMZj4gmS8qZy9cUhyYIcuiEz8KQO2GazBqfjsGIRVyIS1vHEhFOjUKfaDTbaUWZ3rQdfbsackptiHg3vD6O678QCL94eD898EgmV/jxYLv79VJvcBrFMjoqlTl3Ftz7kOYsQigbNG6u7c4pOWXyzOxWN1z7oHSn64SM+kdKsh/IbwLj/LkdCVLlZwcSmj3iAv8vmTVRzXZP0CI6E5gvqx9QfYS21Jitd0hCNDodKlEbCbt0kEiHwEpl6DIuMtN/4NZ/JLOCpV1fciqdaslYYyo53ZvIf2NwoSNU4ie6P4iREuvrBjh6sMBk3ovDk5gTxSmAPAgfLPOWDgrUupCOUoqLGVUpsJov0ArfKJ/iJsUvr0psuZPE2FHhZ5n7U6f2EhWFvaqv5LaBw3EcxwtFd86wCh/AYCFNSHi73UEy2QOXNTmXvxTc9dwVSzmeF1eWujwkzpi7QR16k2smAUb67/U51Auw2B47Hm9/Atshx4msmTw83Us8csoDdRf382+fOOo+xK6OHtBfpQSW+ZHGnr1FQcXZrZ0Q2xSBJ8xAMSRgdtmVhPuJXGwcw3qPnV/uYHj5hnXudMFGf1w5K+soOSRsKZ3SNom4NlszzdNuMBvKpUzKtlsNKmAQb8zuVTo0vnhFrdfAokiUCyvHs8fX9r0T79rB2C+3iV+F1DX/Cp3BB8HjqH10ZsH/o/8s59mb7GxWK76zVI1hfkMjXwYir6AoqqcSM761iHigeCLXBuqpjQLwKvyLQlsjM/vz5Cb5DamqCDyB6Y3JsbWHrC+8DU3KEtL9XtYElKj5MkEixJPMcJcSLrs=");
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
                var children = doc.DocumentNode.SelectNodes("//span[contains(@id,'D539Control_history1_dlQuery_D539_DDate')] | //span[contains(@id,'D539Control_history1_dlQuery_SNo')]");

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
