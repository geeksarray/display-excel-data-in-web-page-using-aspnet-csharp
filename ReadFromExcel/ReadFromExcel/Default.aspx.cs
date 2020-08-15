using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;

namespace ReadFromExcel
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                OleDbConnection oconn = null;
                string FilePath = "C:\geeksarray\Products.xlsx";
                oconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties=Excel 8.0");

                DataTable dtCategories = new DataTable();
                oconn.Open();
                dtCategories = oconn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                oconn.Close();

                List<ListItem> lstCategories = new List<ListItem>();

                string sheetName = string.Empty;
                foreach (DataRow dr in dtCategories.Rows)
                {
                    sheetName = dr["TABLE_NAME"].ToString();
                    if (!sheetName.Contains("Sheet"))
                        lstCategories.Add(new ListItem(sheetName.Replace("$", "")));
                }
                ddlCategories.DataSource = lstCategories;
                ddlCategories.DataBind();
            }
        }

        protected void ddlCategories_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dtProducts = new DataTable();
            dtProducts.Columns.Add("CategoryName");
            dtProducts.Columns.Add("ProductID");
            dtProducts.Columns.Add("ProductName");
            dtProducts.Columns.Add("SupplierID");
            dtProducts.Columns.Add("QuantityPerUnit");
            dtProducts.Columns.Add("UnitPrice");
            dtProducts.Columns.Add("UnitsInStock");
            dtProducts.Columns.Add("UnitsOnOrder");
            dtProducts.Columns.Add("ReorderLevel");
            dtProducts.Columns.Add("Discontinued");

            OleDbConnection oconn = null;
            string FilePath = "C:\geeksarray\Products.xlsx";
            oconn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties=Excel 8.0");
            OleDbCommand ocmd = new OleDbCommand("select * from [" + ddlCategories.SelectedItem.Value + "$]", oconn);

            oconn.Open();
            OleDbDataReader odr = ocmd.ExecuteReader();

            while (odr.Read())
            {
                DataRow drProducts = dtProducts.NewRow();
                drProducts["CategoryName"] = HandleNull(odr.GetValue(0));
                drProducts["ProductID"] = HandleNull(odr.GetValue(1));
                drProducts["ProductName"] = HandleNull(odr.GetValue(2));
                drProducts["SupplierID"] = HandleNull(odr.GetValue(3));
                drProducts["QuantityPerUnit"] = HandleNull(odr.GetValue(4));
                drProducts["UnitPrice"] = HandleNull(odr.GetValue(5));
                drProducts["UnitsInStock"] = HandleNull(odr.GetValue(6));
                drProducts["UnitsOnOrder"] = HandleNull(odr.GetValue(7));
                drProducts["ReorderLevel"] = HandleNull(odr.GetValue(8));
                drProducts["Discontinued"] = HandleNull(odr.GetValue(9));
                dtProducts.Rows.Add(drProducts);  
            }

            oconn.Close();

            gvProducts.DataSource = dtProducts;
            gvProducts.DataBind();  
        }

        private string HandleNull(object val)
        {
            if (val == null)
                return string.Empty;

            return val.ToString(); 
        }
    }
}
