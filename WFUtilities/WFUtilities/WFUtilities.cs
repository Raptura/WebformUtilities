using System;
using System.Data;
using System.IO;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.DirectoryServices;
using System.Web.Services;
using System.Runtime.Serialization.Formatters.Binary;
using System.Web.UI.HtmlControls;
using System.Windows.Forms;


/*************************************************************************************************
* Author: Armond Smith
* Created On: 5/16/2016
* 
* Last Modified By:
* Last Modified On:
* 
* Authorized Contributors:
*
* Version 1.0.0
**************************************************************************************************/

namespace WFUtilities
{
    public class WFUtilities
    {


        /*************************************************************************************************
        * 
        * Excell Values
        * 
        **************************************************************************************************/

        //GridViews
        public static void GridViewToXLS(GridView gv, HttpResponse response, string fileName)
        {
            response.Clear();
            response.AddHeader("content-disposition", "attachment;filename=" + fileName + ".xls");
            response.Charset = "";
            response.ContentType = "application/vnd.ms-excel";
            //Prevent grid splitting

            GridView newGv = new GridView();
            newGv.DataSource = gv.DataSource;
            newGv.DataBind();

            StringWriter StringWriter = new StringWriter();
            HtmlTextWriter HtmlTextWriter = new HtmlTextWriter(StringWriter);

            //Prevent grid splitting
            newGv.AllowPaging = false;
            newGv.AllowSorting = false;

            newGv.RenderControl(HtmlTextWriter);
            response.Write(StringWriter.ToString());
            response.Flush();
            response.End();
        }

        public static void GridViewToXLS(GridView gv, HttpResponse response, string fileName, int startCol)
        {
            DataSet myDataSet = new DataSet();
            myDataSet = (DataSet)gv.DataSource;
            DataTable myTable = new DataTable(gv.ToString());

            for (int i = startCol; i < myTable.Columns.Count; i++)
            {
                myTable.Columns.Add(myDataSet.Tables[0].Columns[i].ToString());

            }
            for (int j = 0; j < myDataSet.Tables[0].Rows.Count; j++)
            {
                int length = myDataSet.Tables[0].Columns.Count - startCol;
                string[] colValues = new string[length];

                for (int k = 0; k < length; k++)
                {
                    colValues[k] = myDataSet.Tables[0].Rows[j][k].ToString();
                }
                myTable.Rows.Add(colValues);
            }


            GridView exportGv = new GridView();
            exportGv.AllowPaging = false;
            exportGv.AllowSorting = false;
            exportGv.DataSource = myTable;
            exportGv.DataBind();

            GridViewToXLS(exportGv, response, fileName);
        }

        public static void GridViewToXLS(GridView gv, HttpResponse response, string fileName, int startCol, int endCol)
        {
            DataSet myDataSet = new DataSet();
            myDataSet = (DataSet)gv.DataSource;
            DataTable myTable = new DataTable(gv.ToString());

            for (int i = startCol; i < endCol; i++)
            {
                myTable.Columns.Add(myDataSet.Tables[0].Columns[i].ToString());

            }
            for (int j = 0; j < myDataSet.Tables[0].Rows.Count; j++)
            {
                int length = endCol - startCol;
                string[] colValues = new string[length];

                for (int k = 0; k < length; k++)
                {
                    colValues[k] = myDataSet.Tables[0].Rows[j][k].ToString();
                }
                myTable.Rows.Add(colValues);
            }


            GridView exportGv = new GridView();
            exportGv.AllowPaging = false;
            exportGv.AllowSorting = false;
            exportGv.DataSource = myTable;
            exportGv.DataBind();

            GridViewToXLS(exportGv, response, fileName);
        }

        //DataTables (Relies on GridView Functions)
        public static void DataTableToXLS(DataTable dt, HttpResponse response, string fileName)
        {
            GridView gv = new GridView();
            gv.DataSource = dt;
            gv.DataBind();

            GridViewToXLS(gv, response, fileName);
        }

        public static void DataTableToXLS(DataTable dt, HttpResponse response, string fileName, int startCol)
        {
            DataTable myTable = dt.Copy();

            for (int i = startCol; i < myTable.Columns.Count; i++)
            {
                myTable.Columns.Add(dt.Columns[i].ToString());

            }
            for (int j = 0; j < dt.Rows.Count; j++)
            {
                int length = dt.Columns.Count - startCol;
                string[] colValues = new string[length];

                for (int k = 0; k < length; k++)
                {
                    colValues[k] = dt.Rows[j][k].ToString();
                }
                myTable.Rows.Add(colValues);
            }
            myTable.AcceptChanges();

            DataTableToXLS(myTable, response, fileName);
        }

        public static void DataTableToXLS(DataTable dt, HttpResponse response, string fileName, int startCol, int endCol)
        {
            DataTable myTable = dt.Copy();

            for (int i = startCol; i < endCol; i++)
            {
                myTable.Columns.Add(dt.Columns[i].ToString());

            }
            for (int j = 0; j < dt.Rows.Count; j++)
            {
                int length = endCol - startCol;
                string[] colValues = new string[length];

                for (int k = 0; k < length; k++)
                {
                    colValues[k] = dt.Rows[j][k].ToString();
                }
                myTable.Rows.Add(colValues);
            }
            myTable.AcceptChanges();

            DataTableToXLS(myTable, response, fileName);
        }

        //Set to Generic
        public static DataTable XLSToDataTable(string inputFile)
        {
            throw new NotImplementedException();
        }

        public static string[][] XLSToArray(string inputFile)
        {
            throw new NotImplementedException();
        }

        /*************************************************************************************************
         * 
         * Comma Separated Values
         * 
        **************************************************************************************************/

        //GridViews
        public static void GridViewToCSV(GridView gv, HttpResponse response, string fileName)
        {
            response.Clear();
            response.AddHeader("content-disposition", "attachment;filename=" + fileName + ".csv");
            response.Charset = "";
            response.ContentType = "application/CSV";
            //Prevent grid splitting

            DataTable dt = gv.DataSource as DataTable;

            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int k = 0; k < dt.Columns.Count; k++)
                {
                    //add separator
                    sb.Append(dt.Rows[i].ItemArray[k].ToString() + ',');
                }
                if (i < dt.Rows.Count - 1)
                    //append new line
                    sb.Append("\r\n");
            }


            response.Output.Write(sb.ToString());
            response.Flush();
            response.End();
        }

        public static void GridViewToCSV(GridView gv, HttpResponse response, string fileName, int startCol)
        {
            DataSet myDataSet = new DataSet();
            myDataSet = (DataSet)gv.DataSource;
            DataTable myTable = new DataTable(gv.ToString());

            for (int i = startCol; i < myTable.Columns.Count; i++)
            {
                myTable.Columns.Add(myDataSet.Tables[0].Columns[i].ToString());

            }
            for (int j = 0; j < myDataSet.Tables[0].Rows.Count; j++)
            {
                int length = myDataSet.Tables[0].Columns.Count - startCol;
                string[] colValues = new string[length];

                for (int k = 0; k < length; k++)
                {
                    colValues[k] = myDataSet.Tables[0].Rows[j][k].ToString();
                }
                myTable.Rows.Add(colValues);
            }


            GridView exportGv = new GridView();
            exportGv.AllowPaging = false;
            exportGv.AllowSorting = false;
            exportGv.DataSource = myTable;
            exportGv.DataBind();

            GridViewToCSV(exportGv, response, fileName);
        }

        public static void GridViewToCSV(GridView gv, HttpResponse response, string fileName, int startCol, int endCol)
        {
            DataSet myDataSet = new DataSet();
            myDataSet = (DataSet)gv.DataSource;
            DataTable myTable = new DataTable(gv.ToString());

            for (int i = startCol; i < endCol; i++)
            {
                myTable.Columns.Add(myDataSet.Tables[0].Columns[i].ToString());

            }
            for (int j = 0; j < myDataSet.Tables[0].Rows.Count; j++)
            {
                int length = endCol - startCol;
                string[] colValues = new string[length];

                for (int k = 0; k < length; k++)
                {
                    colValues[k] = myDataSet.Tables[0].Rows[j][k].ToString();
                }
                myTable.Rows.Add(colValues);
            }


            GridView exportGv = new GridView();
            exportGv.AllowPaging = false;
            exportGv.AllowSorting = false;
            exportGv.DataSource = myTable;
            exportGv.DataBind();

            GridViewToCSV(exportGv, response, fileName);
        }

        //DataTables
        public static void DataTableToCSV(DataTable dt, HttpResponse response, string fileName)
        {
            response.Clear();
            response.AddHeader("content-disposition", "attachment;filename=" + fileName + ".csv");
            response.Charset = "";
            response.ContentType = "application/CSV";

            StringBuilder sb = new StringBuilder();

            //append new line
            sb.Append("\r\n");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int k = 0; k < dt.Columns.Count; k++)
                {
                    //add separator
                    sb.Append(dt.Rows[i].ItemArray[k].ToString() + ',');
                }
                //append new line
                sb.Append("\r\n");
            }


            response.Output.Write(sb.ToString());
            response.Flush();
            response.End();
        }

        public static void DataTableToCSV(DataTable dt, HttpResponse response, string fileName, int startCol)
        {
            DataTable myTable = dt.Copy();

            for (int i = startCol; i < myTable.Columns.Count; i++)
            {
                myTable.Columns.Add(dt.Columns[i].ToString());

            }
            for (int j = 0; j < dt.Rows.Count; j++)
            {
                int length = dt.Columns.Count - startCol;
                string[] colValues = new string[length];

                for (int k = 0; k < length; k++)
                {
                    colValues[k] = dt.Rows[j][k].ToString();
                }
                myTable.Rows.Add(colValues);
            }
            myTable.AcceptChanges();

            DataTableToCSV(myTable, response, fileName);
        }

        public static void DataTableToCSV(DataTable dt, HttpResponse response, string fileName, int startCol, int endCol)
        {
            DataTable myTable = dt.Copy();

            for (int i = startCol; i < endCol; i++)
            {
                myTable.Columns.Add(dt.Columns[i].ToString());

            }
            for (int j = 0; j < dt.Rows.Count; j++)
            {
                int length = endCol - startCol;
                string[] colValues = new string[length];

                for (int k = 0; k < length; k++)
                {
                    colValues[k] = dt.Rows[j][k].ToString();
                }
                myTable.Rows.Add(colValues);
            }
            myTable.AcceptChanges();

            DataTableToCSV(myTable, response, fileName);
        }


        public static DataTable CSVToDataTable(string inputFile)
        {
            DataTable dt = new DataTable();
            FileStream fs = File.OpenRead(inputFile);
            StreamReader sr = new StreamReader(fs);
            string totalContent = sr.ReadToEnd();
            string[] lines = totalContent.Split("\r\n"[0]);

            for (int i = 0; i < lines.Length; i++)
            {
                dt.Rows.Add();
                string[] lineContent = lines[i].Split(","[0]); //should be the same length as the colLength
                for (int j = 0; j < lineContent.Length; j++)
                {
                    if (dt.Columns.Count < j)
                    {
                        dt.Columns.Add();
                    }
                    dt.Rows[i][j] = lineContent[j];
                }
            }

            dt.AcceptChanges();
            return dt;
        }

        public static string[][] CSVToArray(string inputFile)
        {
            string[][] arr;
            FileStream fs = File.OpenRead(inputFile);
            StreamReader sr = new StreamReader(fs);
            string totalContent = sr.ReadToEnd();
            string[] lines = totalContent.Split("\r\n"[0]);
            int colLength = lines[0].Split(","[0]).Length;

            arr = new string[lines.Length][];

            for (int i = 0; i < lines.Length; i++)
            {
                arr[i] = new string[colLength];
                string[] lineContent = lines[i].Split(","[0]); //should be the same length as the colLength
                for (int j = 0; j < lineContent.Length; j++)
                {
                    arr[i][j] = lineContent[j];
                }
            }

            return arr;
        }


        /*************************************************************************************************
         * 
         * PDF Serialization
         * 
        **************************************************************************************************/

        [Obsolete]
        public static byte[] SerializePDF(FileUpload control, string fileName)
        {
            Boolean fileOK = false;

            if (control.HasFile)
            {
                string fileExtention = System.IO.Path.GetExtension(control.FileName).ToLower();
                string[] allowedExtentions = { ".pdf" };
                for (int i = 0; i < allowedExtentions.Length; i++)
                {
                    if (fileExtention == allowedExtentions[i])
                    {
                        fileOK = true;
                    }
                }
            }

            if (fileOK)
            {
                String path = System.Web.Hosting.HostingEnvironment.MapPath("~/UploadedForms/");
                control.PostedFile.SaveAs(path + fileName);
                string filePath = path + fileName;

                //FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                //byte[] bytes = new byte[fs.Length];
                byte[] bytes = System.IO.File.ReadAllBytes(filePath);
                //fs.Read(bytes, 0, (int)fs.Length);
                //NARF_DB.FirstUpload(bytes);

                //Put it in the SQL server
                return bytes;
            }
            else
            {
                //Put a modal here that says that the file type must be a PDF file
                return null;
            }
        }

        [Obsolete]
        public static void DeserializePDF(HttpResponse Response, string fileName, byte[] pdfData)
        {
            Response.Clear();
            Response.AddHeader("content-disposition",
            "attachment;filename=" + fileName + ".pdf");
            Response.Charset = "";
            Response.ContentType = "application/pdf";

            Response.Buffer = true;

            //convert to the pdf
            MemoryStream ms = new MemoryStream(pdfData);
            ms.WriteTo(Response.OutputStream);
            Response.End();

            //System.IO.File.WriteAllBytes(fileName, pdfData);
        }

    }
}
