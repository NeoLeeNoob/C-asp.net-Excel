using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.OleDb;
using System.Data;

namespace ECS
{
    public partial class ECS : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void Button1_Click(object sender, EventArgs e)
        {
            if (FileUpload1.HasFile)
            {
                string fileName = FileUpload1.FileName;
                string filePath = Server.MapPath("~/App_Data") + "\\" + fileName;

                // 讀取特定工作表的資料
                string sheetName = "工作表1";

                using (OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=Excel 12.0 XML;"))
                {
                    conn.Open();

                    string sql = "SELECT * FROM [" + sheetName + "$]";
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        using (OleDbDataReader reader = cmd.ExecuteReader())
                        {
                            DataTable dt = new DataTable();
                            dt.Load(reader);

                            // 自訂 GridView 控制項的欄位
                            GridView1.AutoGenerateColumns = false;

                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                BoundField field = new BoundField();
                                field.DataField = reader.GetName(i);
                                field.HeaderText = reader.GetName(i);

                                GridView1.Columns.Add(field);
                            }

                            GridView1.DataSource = dt;
                            GridView1.DataBind();
                        }
                    }
                }
            }
        }

    }
}