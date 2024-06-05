using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;
using System.Data.Common;
using System.Diagnostics.Eventing.Reader;
using System.Reflection.Emit;
using System.Xml.Linq;
using System.Windows.Input;
using static OfficeOpenXml.ExcelErrorValue;

namespace ECS
{
    public partial class ECS2 : System.Web.UI.Page
    {
        string connetionString = "provider=Microsoft.ACE.OLEDB.12.0; Data Source=D:\\ECS\\ECS\\ECS\\ECS\\App_Data\\個人資料1.xlsx ; Extended Properties = \"Excel 12.0 Xml;HDR=yes; IMEX = 1\" ";

        protected void CheckBox1_CheckedChanged1(object sender, EventArgs e)//勾選顯示隱藏......................................................................................................................................
        {
            GridView1.Columns[2].Visible = CheckBox1.Checked;
            if (CheckBox1.Checked == false)
            {
                for (int i = 1; i < GridView1.Rows.Count; i++)
                { GridView1.Rows[i].Cells[2].Visible = false; }
            }
            else
            {
                for (int i = 1; i < GridView1.Rows.Count; i++)
                { GridView1.Rows[i].Cells[2].Visible = true; }
            }
        }
        protected void CheckBox2_CheckedChanged2(object sender, EventArgs e)//勾選顯示隱藏......................................................................................................................................
        {
            GridView1.Columns[3].Visible = CheckBox2.Checked;
            if (CheckBox2.Checked == false)
            {
                for (int i = 2; i < GridView1.Rows.Count; i++)
                { GridView1.Rows[i].Cells[3].Visible = false; }
            }
            else
            {
                for (int i = 2; i < GridView1.Rows.Count; i++)
                { GridView1.Rows[i].Cells[3].Visible = true; }
            }
        }
        protected void CheckBox3_CheckedChanged3(object sender, EventArgs e)//勾選顯示隱藏......................................................................................................................................
        {
            GridView1.Columns[4].Visible = CheckBox3.Checked;
            if (CheckBox3.Checked == false)
            {
                for (int i = 3; i < GridView1.Rows.Count; i++)
                { GridView1.Rows[i].Cells[4].Visible = false; }
            }
            else
            {
                for (int i = 3; i < GridView1.Rows.Count; i++)
                { GridView1.Rows[i].Cells[4].Visible = true; }
            }
        }
        protected void CheckBox4_CheckedChanged4(object sender, EventArgs e)//勾選顯示隱藏......................................................................................................................................
        {
            GridView1.Columns[5].Visible = CheckBox4.Checked;
            if (CheckBox4.Checked == false)
            {
                for (int i = 4; i < GridView1.Rows.Count; i++)
                { GridView1.Rows[i].Cells[5].Visible = false; }
            }
            else
            {
                for (int i = 4; i < GridView1.Rows.Count; i++)
                { GridView1.Rows[i].Cells[5].Visible = true; }
            }
        }
        protected void CheckBox5_CheckedChanged5(object sender, EventArgs e)//勾選顯示隱藏......................................................................................................................................
        {
            GridView1.Columns[6].Visible = CheckBox5.Checked;
            if (CheckBox5.Checked == false)
            {
                for (int i = 5; i < GridView1.Rows.Count; i++)
                { GridView1.Rows[i].Cells[6].Visible = false; }
            }
            else
            {
                for (int i = 5; i < GridView1.Rows.Count; i++)
                { GridView1.Rows[i].Cells[6].Visible = true; }
            }
        }
        protected void Page_Load(object sender, EventArgs e)//網站頁面........................................................................................................................................................................................
        {
            string sql = "SELECT TOP 10 * FROM  [工作表1$] ORDER BY ID DESC ";
            if (!IsPostBack)
            {
                using (OleDbConnection connection = new OleDbConnection(connetionString))
                {
                    using (OleDbDataAdapter dataadapter = new OleDbDataAdapter(sql, connection))
                    {
                        connection.Open();
                        DataSet ds = new DataSet();
                        dataadapter.Fill(ds);
                        GridView1.DataSource = ds.Tables[0];
                        GridView1.DataBind();
                        connection.Close();
                    }
                }
                GridView1.HeaderRow.Cells[1].Visible = false; for (int i = 0; i < GridView1.Rows.Count; i++) { GridView1.Rows[i].Cells[1].Visible = false; }
            };
        }
        protected void ImageButton1_Click(object sender, ImageClickEventArgs e)  //獲取該列資料得已修改按鈕..........................................................................................................
        {
            // 獲取 ImageButton 所在的資料列
            GridViewRow row = (GridViewRow)((ImageButton)sender).NamingContainer;
            // 獲取要修改的資料列的資料
            string swpa = row.Cells[1].Text;
            string value1 = row.Cells[2].Text;
            string value2 = row.Cells[3].Text;
            string value3 = row.Cells[4].Text;
            string value4 = row.Cells[5].Text;
            string value5 = row.Cells[6].Text;
            TextBox1.Text = value1;
            TextBox2.Text = value2;
            TextBox3.Text = value3;
            TextBox4.Text = value4;
            TextBox5.Text = value5;
            ViewState["swpa"] = swpa;
        }
        protected void ImageButton2_Click(object sender, ImageClickEventArgs e)  //刪除按鈕.............................................................................................................................................
        {
            // 獲取 ImageButton 所在的資料列
            GridViewRow row = (GridViewRow)((ImageButton)sender).NamingContainer;
            // 獲取要修改的資料列的資料
            string DELID = row.Cells[1].Text;
            string sql = "UPDATE [工作表1$] SET ID = Null, 名字 = Null, 公司 = Null, 料號 = Null, 地址 = Null, 料件 = Null WHERE ID = " + DELID + "";
            using (OleDbConnection connection = new OleDbConnection(connetionString))
            {
                using (OleDbCommand command = new OleDbCommand(sql, connection))
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();
                }
            }
            // 重新讀取資料並更新 GridView1
            sql = "SELECT TOP 10 * FROM  [工作表1$] WHERE 名字 IS NOT NULL AND 公司 IS NOT NULL ORDER BY ID DESC  ";
            using (OleDbConnection connection = new OleDbConnection(connetionString))
            {
                using (OleDbDataAdapter dataadapter = new OleDbDataAdapter(sql, connection))
                {
                    connection.Open();
                    DataSet ds = new DataSet();
                    dataadapter.Fill(ds);
                    GridView1.DataSource = ds.Tables[0];
                    GridView1.DataBind();
                    connection.Close();
                }
            }
            GridView1.HeaderRow.Cells[1].Visible = false;
            for (int i = 0; i < GridView1.Rows.Count; i++) { GridView1.Rows[i].Cells[1].Visible = false; }
        }
        protected void Button1_Click(object sender, EventArgs e)  //新增按鈕................................................................................................................................................................................
        {
            if (TextBox1.Text == "")
            {
                // 如果 TextBox1.Text 值為 null，顯示錯誤訊息
                Label1.Text = "請先填寫完整要新增的資料。";
                return;
            }
            else if (TextBox2.Text == "")
            {
                // 如果 TextBox2.Text 值為 null，顯示錯誤訊息
                Label1.Text = "請先填寫完整要新增的資料。";
                return;
            }
            String value1 = TextBox1.Text;
            String value2 = TextBox2.Text;
            String value3 = TextBox3.Text;
            String value4 = TextBox4.Text;
            String value5 = TextBox5.Text;
            //先抓取最大的ID
            int ID = 0;
            String sql = "SELECT MAX(ID) FROM [工作表1$] ";

            using (OleDbConnection connection = new OleDbConnection(connetionString))
            {
                using (OleDbCommand command = new OleDbCommand(sql, connection))
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        ID = Convert.ToInt32(reader[0]);
                    }
                    connection.Close();
                }
            }
            //再新增一筆資料
            ID = ID + 1;
            sql = "INSERT INTO [工作表1$] (ID, 名字, 公司, 料號, 地址, 料件) VALUES ('" + ID + "','" + value1 + "','" + value2 + "','" + value3 + "','" + value4 + "','" + value5 + "')";
            TextBox1.Text = "";
            TextBox2.Text = "";
            TextBox3.Text = "";
            TextBox4.Text = "";
            TextBox5.Text = "";
            Label1.Text += ID + sql;

            using (OleDbConnection connection = new OleDbConnection(connetionString))
            {
                using (OleDbCommand command = new OleDbCommand(sql, connection))
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();
                }
            }
          /*  Response.Redirect(Request.Url.AbsoluteUri);
            GridView1.HeaderRow.Cells[1].Visible = false;
            for (int i = 0; i < GridView1.Rows.Count; i++) { GridView1.Rows[i].Cells[1].Visible = false; }*/
        }
        protected void Button2_Click(object sender, EventArgs e)  //修改按鈕................................................................................................................................................................................
        {
            // 檢查 ViewState 中的 swpa 值是否為 null
            if (ViewState["swpa"] == null)
            {
                // 如果 swpa 值為 null，顯示錯誤訊息
                Label1.Text = "請先選擇要修改的資料列。";
                return;
            }
            // 獲取要修改的資料列的資料
            string swpa = (string)ViewState["swpa"];
            string upvalue1 = TextBox1.Text;
            string upvalue2 = TextBox2.Text;
            string upvalue3 = TextBox3.Text;
            string upvalue4 = TextBox4.Text;
            string upvalue5 = TextBox5.Text;

            // 更新資料庫中的資料
            string sql = "UPDATE [工作表1$] SET 名字 = @upvalue1, 公司 = @upvalue2, 料號 = @upvalue3, 地址 = @upvalue4, 料件 = @upvalue5 WHERE ID = @swpa";
            using (OleDbConnection connection = new OleDbConnection(connetionString))
            {
                using (OleDbCommand command = new OleDbCommand(sql, connection))
                {
                    command.Parameters.AddWithValue("@upvalue1", upvalue1);
                    command.Parameters.AddWithValue("@upvalue2", upvalue2);
                    command.Parameters.AddWithValue("@upvalue3", upvalue3);
                    command.Parameters.AddWithValue("@upvalue4", upvalue4);
                    command.Parameters.AddWithValue("@upvalue5", upvalue5);
                    command.Parameters.AddWithValue("@swpa", swpa);
                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();
                }
            }
            // 將 ViewState 中的 swpa 值設為 null
            ViewState["swpa"] = null;
            TextBox1.Text = "";
            TextBox2.Text = "";
            TextBox3.Text = "";
            TextBox4.Text = "";
            TextBox5.Text = "";
            // 重新讀取資料並更新 GridView1
            sql = "SELECT TOP 10 * FROM  [工作表1$] WHERE 名字 IS NOT NULL AND 公司 IS NOT NULL ORDER BY ID DESC";
            using (OleDbConnection connection = new OleDbConnection(connetionString))
            {
                using (OleDbDataAdapter dataadapter = new OleDbDataAdapter(sql, connection))
                {
                    connection.Open();
                    DataSet ds = new DataSet();
                    dataadapter.Fill(ds);
                    GridView1.DataSource = ds.Tables[0];
                    GridView1.DataBind();
                    connection.Close();
                }
            }
            GridView1.HeaderRow.Cells[1].Visible = false;
            for (int i = 0; i < GridView1.Rows.Count; i++)
            {
                GridView1.Rows[i].Cells[1].Visible = false;
            }
            Response.Redirect(Request.Url.AbsoluteUri);
        }
        protected void Button3_Click(object sender, EventArgs e)  //查詢按鈕.................................................................................................................................................................................
        {
            String value1 = TextBox1.Text;
            String value2 = TextBox2.Text;
            String value3 = TextBox3.Text;
            String value4 = TextBox4.Text;
            String value5 = TextBox5.Text;
            String sql = "SELECT TOP 10 * FROM  [工作表1$] WHERE 名字 IS NOT NULL AND 公司 IS NOT NULL ORDER BY ID DESC";
            if (!String.IsNullOrEmpty(value1) || !String.IsNullOrEmpty(value2) || !String.IsNullOrEmpty(value3) || !String.IsNullOrEmpty(value4) || !String.IsNullOrEmpty(value5))
            {
                sql = "SELECT * FROM [工作表1$] WHERE 1 = 1";
                if (!String.IsNullOrEmpty(value1)) { sql += "AND (名字 IS NULL OR 名字 LIKE @value1)"; }
                if (!String.IsNullOrEmpty(value2)) { sql += "AND (公司 IS NULL OR 公司 LIKE @value2)"; }
                if (!String.IsNullOrEmpty(value3)) { sql += "AND (料號 IS NULL OR 料號 LIKE @value3)"; }
                if (!String.IsNullOrEmpty(value4)) { sql += "AND (地址 IS NULL OR 地址 LIKE @value4)"; }
                if (!String.IsNullOrEmpty(value5)) { sql += "AND (料件 IS NULL OR 料件 LIKE @value5)"; }
                sql += "AND NOT (名字 IS NULL AND 公司 IS NULL AND 料號 IS NULL AND 地址 IS NULL AND 料件 IS NULL)";

            }
            TextBox1.Text = "";
            TextBox2.Text = "";
            TextBox3.Text = "";
            TextBox4.Text = "";
            TextBox5.Text = "";
            using (OleDbConnection connection = new OleDbConnection(connetionString))
            {
                using (OleDbCommand command = new OleDbCommand(sql, connection))
                {
                    if (!String.IsNullOrEmpty(value1)) { command.Parameters.AddWithValue("@value1", "%" + value1 + "%"); }
                    if (!String.IsNullOrEmpty(value2)) { command.Parameters.AddWithValue("@value2", "%" + value2 + "%"); }
                    if (!String.IsNullOrEmpty(value3)) { command.Parameters.AddWithValue("@value3", "%" + value3 + "%"); }
                    if (!String.IsNullOrEmpty(value4)) { command.Parameters.AddWithValue("@value4", "%" + value4 + "%"); }
                    if (!String.IsNullOrEmpty(value5)) { command.Parameters.AddWithValue("@value5", "%" + value5 + "%"); }
                    sql += "AND NOT (名字 IS NULL AND 公司 IS NULL AND 料號 IS NULL AND 地址 IS NULL AND 料件 IS NULL)";
                    connection.Open();
                    DataSet ds = new DataSet();
                    OleDbDataAdapter dataadapter = new OleDbDataAdapter(command);
                    dataadapter.Fill(ds);
                    GridView1.DataSource = ds.Tables[0];
                    GridView1.DataBind();
                    connection.Close();
                }
            }
            GridView1.HeaderRow.Cells[1].Visible = false;
            for (int i = 0; i < GridView1.Rows.Count; i++) { GridView1.Rows[i].Cells[1].Visible = false; }
        }
    }
}