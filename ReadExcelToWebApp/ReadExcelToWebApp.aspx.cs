using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Configuration;
using System.Globalization;

namespace ReadExtelToEb
{
    public partial class ReadExcelToWeb : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            var filePath = ConfigurationManager.AppSettings["FilePath"];
            string ext = Path.GetExtension(filePath);
            ReadExcel(filePath, ext);
        }

        protected void ReadExcel(string fileName, string fileExt)
        {
            //Coneection String by default empty  
            string ConStr = "";
            //there are two types of extation .xls and .xlsx of Excel   
           // Label1.Text = FileUpload1.FileName + "\'s Data showing into the GridView";
            //checking that extantion is .xls or .xlsx  
            if (fileExt.Trim() == ".xls")
            {
                //connection string for that file which extantion is .xls  
                ConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
            }
            else if (fileExt.Trim() == ".xlsx")
            {
                //connection string for that file which extantion is .xlsx  
                ConStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            }
            //making query  
            string query = "SELECT * FROM [LinkedIn-Posts$]";
            //Providing connection  
            OleDbConnection conn = new OleDbConnection(ConStr);
            //checking that connection state is closed or not if closed the   
            //open the connection  
            if (conn.State == ConnectionState.Closed)
            {
                conn.Open();
            }
            //create command object  
            OleDbCommand cmd = new OleDbCommand(query, conn);
            // create a data adapter and get the data into dataadapter  
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();
            //fill the Excel data to data set  
            da.Fill(ds);
           //Obtained the date from datetime column 
            DataTable dtCloned = ds.Tables[0].Clone();
            dtCloned.Columns[0].DataType = typeof(string);
            for(int i=0;i<ds.Tables[0].Rows.Count;i++)
            {
                DataRow row = ds.Tables[0].Rows[i];
                DateTime dt = DateTime.Parse(row["Date"].ToString());
                dtCloned.ImportRow(row);
                dtCloned.Rows[i]["Date"]= dt.ToShortDateString();
            }

            //set data source of the grid view           
            gvExcelFile.DataSource = dtCloned;
            //binding the gridview  
            gvExcelFile.DataBind();
            //close the connection  
            conn.Close();
        }
     


    }
}