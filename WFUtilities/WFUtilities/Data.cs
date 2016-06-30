using System;
using System.Data;
using System.IO;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.Office.Interop.Excel;


/*************************************************************************************************
* Author: Armond Smith
* Created On: 5/16/2016
* 
* Last Modified By: Kyler Love
* Last Modified On: 6/21/2016
* 
* Authorized Contributors:
* Kyler Love
* Version 1.0.6
**************************************************************************************************/

namespace WFUtilities
{
    public class Data
    {

        /// <summary>
        /// Processes Exports of data to Downloadable Files
        /// </summary>
        public class Export
        {
            /*************************************************************************************************
            * 
            * Excell Values
            * 
            **************************************************************************************************/

            //GridViews

            /// <summary>
            /// Exports a Gridview to an Excel File.
            /// </summary>
            /// <param name="gv">The Grid View.</param>
            /// <param name="response">The HttpResponse that will download the Excel File.</param>
            /// <param name="fileName">Name of the file.</param>
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

            /// <summary>
            /// Exports a Gridview to an Excel File.
            /// </summary>
            /// <param name="gv">The Grid View.</param>
            /// <param name="response">The HttpResponse that will download the Excel File.</param>
            /// <param name="fileName">Name of the file.</param>
            /// <param name="startCol">The first column of the Gridview to export.</param>
            /// <exception cref="System.ArgumentOutOfRangeException">Starting column must be within the range of the column count</exception>
            public static void GridViewToXLS(GridView gv, HttpResponse response, string fileName, int startCol)
            {
                DataSet myDataSet = new DataSet();
                myDataSet = (DataSet)gv.DataSource;
                System.Data.DataTable myTable = new System.Data.DataTable(gv.ToString());

                if (startCol > myTable.Columns.Count || startCol < 0)
                    throw new ArgumentOutOfRangeException("Starting column must be within the range of the column count");

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

            /// <summary>
            /// Exports a Gridview to an Excel File.
            /// </summary>
            /// <param name="gv">The Grid View.</param>
            /// <param name="response">The HttpResponse that will download the Excel File.</param>
            /// <param name="fileName">Name of the file.</param>
            /// <param name="startCol">The first column of the Gridview to export.</param>
            /// <param name="endCol">The last column of the Gridview to export.</param>
            /// <exception cref="System.ArgumentOutOfRangeException">
            /// Starting column must be within the range of the column count
            /// or
            /// Starting Column must be before the ending column
            /// or
            /// Ending Column must be within range of the column count
            /// </exception>
            public static void GridViewToXLS(GridView gv, HttpResponse response, string fileName, int startCol, int endCol)
            {
                DataSet myDataSet = new DataSet();
                myDataSet = (DataSet)gv.DataSource;
                System.Data.DataTable myTable = new System.Data.DataTable(gv.ToString());

                if (startCol > myTable.Columns.Count || startCol < 0)
                    throw new ArgumentOutOfRangeException("Starting column must be within the range of the column count");
                if (startCol > endCol)
                    throw new ArgumentOutOfRangeException("Starting Column must be before the ending column");
                if (endCol > myTable.Columns.Count || endCol < 0)
                    throw new ArgumentOutOfRangeException("Ending Column must be within range of the column count");


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

            /// <summary>
            /// Export a DataTable to an Excel File
            /// </summary>
            /// <param name="dt">The Data Table.</param>
            /// <param name="response">The HttpResponse that will download the Excel File.</param>
            /// <param name="fileName">Name of the file.</param>
            public static void DataTableToXLS(System.Data.DataTable dt, HttpResponse response, string fileName)
            {
                GridView gv = new GridView();
                gv.DataSource = dt;
                gv.DataBind();

                GridViewToXLS(gv, response, fileName);
            }

            /// <summary>
            /// Export a DataTable to an Excel File
            /// </summary>
            /// <param name="dt">The Data Table.</param>
            /// <param name="response">The HttpResponse that will download the Excel File.</param>
            /// <param name="fileName">Name of the file.</param>
            /// <param name="startCol">The first column of the Gridview to export.</param>
            /// <exception cref="System.ArgumentOutOfRangeException">Starting column must be within the range of the column count</exception>
            public static void DataTableToXLS(System.Data.DataTable dt, HttpResponse response, string fileName, int startCol)
            {
                System.Data.DataTable myTable = dt.Copy();

                if (startCol > myTable.Columns.Count || startCol < 0)
                    throw new ArgumentOutOfRangeException("Starting column must be within the range of the column count");

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

            /// <summary>
            /// Export a DataTable to an Excel File
            /// </summary>
            /// <param name="dt">The Data Table.</param>
            /// <param name="response">The HttpResponse that will download the Excel File.</param>
            /// <param name="fileName">Name of the file.</param>
            /// <param name="startCol">The first column of the Gridview to export.</param>
            /// <param name="endCol">The last column of the Gridview to export.</param>
            /// <exception cref="System.ArgumentOutOfRangeException">
            /// Starting column must be within the range of the column count
            /// or
            /// Starting Column must be before the ending column
            /// or
            /// Ending Column must be within range of the column count
            /// </exception>
            public static void DataTableToXLS(System.Data.DataTable dt, HttpResponse response, string fileName, int startCol, int endCol)
            {
                System.Data.DataTable myTable = dt.Copy();

                if (startCol > myTable.Columns.Count || startCol < 0)
                    throw new ArgumentOutOfRangeException("Starting column must be within the range of the column count");
                if (startCol > endCol)
                    throw new ArgumentOutOfRangeException("Starting Column must be before the ending column");
                if (endCol > myTable.Columns.Count || endCol < 0)
                    throw new ArgumentOutOfRangeException("Ending Column must be within range of the column count");


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

            /*************************************************************************************************
            * 
            * Comma Separated Values
            * 
            **************************************************************************************************/

            //GridViews

            /// <summary>
            /// Export a DataTable to an CSV File
            /// </summary>
            /// <param name="gv">The Gridview.</param>
            /// <param name="response">The HttpResponse that will download the Excel File.</param>
            /// <param name="fileName">Name of the file.</param>
            public static void GridViewToCSV(GridView gv, HttpResponse response, string fileName)
            {
                response.Clear();
                response.AddHeader("content-disposition", "attachment;filename=" + fileName + ".csv");
                response.Charset = "";
                response.ContentType = "application/CSV";
                //Prevent grid splitting

                System.Data.DataTable dt = gv.DataSource as System.Data.DataTable;

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

            /// <summary>
            /// Export a DataTable to an CSV File
            /// </summary>
            /// <param name="gv">The Gridview.</param>
            /// <param name="response">The HttpResponse that will download the Excel File.</param>
            /// <param name="fileName">Name of the file.</param>
            /// <param name="startCol">The first column of the Gridview to export.</param>
            /// <exception cref="System.ArgumentOutOfRangeException">Starting column must be within the range of the column count</exception>
            public static void GridViewToCSV(GridView gv, HttpResponse response, string fileName, int startCol)
            {
                DataSet myDataSet = new DataSet();
                myDataSet = (DataSet)gv.DataSource;
                System.Data.DataTable myTable = new System.Data.DataTable(gv.ToString());

                if (startCol > myTable.Columns.Count || startCol < 0)
                    throw new ArgumentOutOfRangeException("Starting column must be within the range of the column count");

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

            /// <summary>
            /// Export a DataTable to an CSV File
            /// </summary>
            /// <param name="gv">The Gridview.</param>
            /// <param name="response">The HttpResponse that will download the Excel File.</param>
            /// <param name="fileName">Name of the file.</param>
            /// <param name="startCol">The first column of the Gridview to export.</param>
            /// <param name="endCol">The last column of the Gridview to export.</param>
            /// <exception cref="System.ArgumentOutOfRangeException">
            /// Starting column must be within the range of the column count
            /// or
            /// Starting Column must be before the ending column
            /// or
            /// Ending Column must be within range of the column count
            /// </exception>
            public static void GridViewToCSV(GridView gv, HttpResponse response, string fileName, int startCol, int endCol)
            {
                DataSet myDataSet = new DataSet();
                myDataSet = (DataSet)gv.DataSource;
                System.Data.DataTable myTable = new System.Data.DataTable(gv.ToString());

                if (startCol > myTable.Columns.Count || startCol < 0)
                    throw new ArgumentOutOfRangeException("Starting column must be within the range of the column count");
                if (startCol > endCol)
                    throw new ArgumentOutOfRangeException("Starting Column must be before the ending column");
                if (endCol > myTable.Columns.Count || endCol < 0)
                    throw new ArgumentOutOfRangeException("Ending Column must be within range of the column count");

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

            /// <summary>
            /// Export a DataTable to an CSV File
            /// </summary>
            /// <param name="dt">The Data Table.</param>
            /// <param name="response">The HttpResponse that will download the Excel File.</param>
            /// <param name="fileName">Name of the file.</param>
            public static void DataTableToCSV(System.Data.DataTable dt, HttpResponse response, string fileName)
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

            /// <summary>
            /// Export a DataTable to an CSV File
            /// </summary>
            /// <param name="dt">The Data Table.</param>
            /// <param name="response">The HttpResponse that will download the Excel File.</param>
            /// <param name="fileName">Name of the file.</param>
            /// <param name="startCol">The first column of the Gridview to export.</param>
            /// <exception cref="System.ArgumentOutOfRangeException">Starting column must be within the range of the column count</exception>
            public static void DataTableToCSV(System.Data.DataTable dt, HttpResponse response, string fileName, int startCol)
            {
                System.Data.DataTable myTable = dt.Copy();

                if (startCol > myTable.Columns.Count || startCol < 0)
                    throw new ArgumentOutOfRangeException("Starting column must be within the range of the column count");

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

            /// <summary>
            /// Export a DataTable to an CSV File
            /// </summary>
            /// <param name="dt">The Data Table.</param>
            /// <param name="response">The HttpResponse that will download the Excel File.</param>
            /// <param name="fileName">Name of the file.</param>
            /// <param name="startCol">The first column of the Gridview to export.</param>
            /// <param name="endCol">The last column of the Gridview to export.</param>
            /// <exception cref="System.ArgumentOutOfRangeException">
            /// Starting column must be within the range of the column count
            /// or
            /// Starting Column must be before the ending column
            /// or
            /// Ending Column must be within range of the column count
            /// </exception>
            public static void DataTableToCSV(System.Data.DataTable dt, HttpResponse response, string fileName, int startCol, int endCol)
            {
                System.Data.DataTable myTable = dt.Copy();

                if (startCol > myTable.Columns.Count || startCol < 0)
                    throw new ArgumentOutOfRangeException("Starting column must be within the range of the column count");
                if (startCol > endCol)
                    throw new ArgumentOutOfRangeException("Starting Column must be before the ending column");
                if (endCol > myTable.Columns.Count || endCol < 0)
                    throw new ArgumentOutOfRangeException("Ending Column must be within range of the column count");

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


        }

        /// <summary>
        /// Imports Files to be used for data manipulation
        /// </summary>
        public class Import
        {
            /// <summary>
            /// Converts a Excel file's data to a Data Table
            /// </summary>
            /// <param name="inputFile">The input file.</param>
            /// <param name="hasHeaderColumn">if set to <c>true</c> [has header column].</param>
            /// <param name="hasHeaderRow">if set to <c>true</c> [has header row].</param>
            /// <param name="columnStart">if set to <c>true</c>[The column start].</param>
            /// <param name="rowStart">if set to <c>true</c>[The row start].</param>
            /// <returns></returns>
            public static System.Data.DataTable XLSToDataTable(string inputFile, bool hasHeaderColumn = false, bool hasHeaderRow = false, int columnStart = 1, int rowStart = 1)
            {
                System.Data.DataTable dt = new System.Data.DataTable();

                Application app = new Application();
                Workbook book = app.Workbooks.Open(@inputFile);
                Worksheet sheet = book.Sheets[1];
                Range range = sheet.UsedRange;

                int m_rowStart = hasHeaderColumn ? rowStart : 0;
                int m_colStart = hasHeaderRow ? columnStart : 0;

                int rowCount = range.Rows.Count - m_rowStart; //effective row count
                int colCount = range.Columns.Count - m_colStart; //effective column count


                for (int i = m_colStart; i < colCount; i++)
                    dt.Columns.Add(hasHeaderColumn ? (sheet.Cells[rowStart, i + 1 + m_colStart].Value as string) : "Column" + i);


                //Set up values
                for (int i = 0; i < rowCount; i++)
                {
                    dt.Rows.Add();
                    for (int j = 0; j < colCount; j++)
                    {
                        dt.Rows[i][j] = sheet.Cells[i + 1 + m_rowStart, j + 1 + m_colStart].Value;
                    }
                }
                app.Quit();
                return dt;
            }

            /// <summary>
            /// Converts a Excel file's data to a string jagged array
            /// </summary>
            /// <param name="inputFile">The input file.</param>
            /// <param name="hasHeaderRow">if set to <c>true</c> [has header row].</param>
            /// <param name="hasHeaderColumn">if set to <c>true</c> [has header column].</param>
            /// <param name="columnStart">if set to <c>true</c> [The column start].</param>
            /// <param name="rowStart">if set to <c>true</c> [The row start].</param>
            /// <returns></returns>
            public static string[][] XLSToArray(string inputFile, bool hasHeaderRow = false, bool hasHeaderColumn = false, int columnStart = 1, int rowStart = 1)
            {
                Application app = new Application();
                Workbook book = app.Workbooks.Open(@inputFile);
                Worksheet sheet = book.Sheets[1];
                Range range = sheet.UsedRange;


                int m_rowStart = hasHeaderColumn ? rowStart : 0;
                int m_colStart = hasHeaderRow ? columnStart : 0;

                int rowCount = range.Rows.Count - m_rowStart; //effective row count
                int colCount = range.Columns.Count - m_colStart; //effective column count

                string[][] output = new string[rowCount][];

                for (int i = 0; i < rowCount; i++)
                {
                    output[i] = new string[colCount];
                    for (int j = 0; j < colCount; j++)
                    {
                        output[i][j] = sheet.Cells[i + 1 + m_rowStart, j + m_colStart + 1].Value;
                    }
                }
                app.Quit();
                return output;
            }


            /// <summary>
            /// Converts a CSV file's data to a Data Table
            /// </summary>
            /// <param name="inputFile">The input file.</param>
            /// <returns></returns>
            public static System.Data.DataTable CSVToDataTable(string inputFile)
            {
                System.Data.DataTable dt = new System.Data.DataTable();
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
                        if (dt.Columns.Count < j + 1)
                        {
                            dt.Columns.Add();
                        }
                        dt.Rows[i][j] = lineContent[j];
                    }
                }

                dt.AcceptChanges();
                return dt;
            }

            /// <summary>
            /// Converts a CSV file's data to a string jagged array
            /// </summary>
            /// <param name="inputFile">The input file.</param>
            /// <returns></returns>
            public static string[][] CSVToArray(string inputFile)
            {
                string[][] arr;
                FileStream fs = File.OpenRead(inputFile);
                StreamReader sr = new StreamReader(fs);
                string totalContent = sr.ReadToEnd();
                string[] lines = totalContent.Split("\r\n"[0]);
                int colLength = lines[0].Split(","[0]).Length;

                arr = new string[lines.Length][];

                for (int i = 0; i < lines.Length; i++) //rows
                {
                    arr[i] = new string[colLength];
                    string[] lineContent = lines[i].Split(","[0]); //should be the same length as the colLength
                    for (int j = 0; j < lineContent.Length; j++) //columns
                    {
                        arr[i][j] = lineContent[j];
                    }
                }

                return arr;
            }


        }

        /// <summary>
        /// Processes Serialization and Deserialization of files into byte data to be stored onto databases
        /// </summary>
        public class Serialization
        {
            /*************************************************************************************************
            * 
            * PDF Serialization
            * 
            **************************************************************************************************/

            /// <summary>
            /// Serializes a file.
            /// </summary>
            /// <param name="control"> The FileUpload control.</param>
            /// <param name="fileName"> Name of the file.</param>
            /// <param name="allowedExtentions"> The allowed extentions.</param>
            /// <param name = "virtualDownloadPath">  The relative path where the file will be stored temporarily</param>
            /// <returns></returns>
            /// <exception cref="System.ArgumentException">Uploaded file must have same extention as allowedExtentions</exception>
            public static byte[] SerializeFile(FileUpload control, string fileName, string[] allowedExtentions, string virtualDownloadPath = "~/UploadedForms/")
            {
                Boolean fileOK = false;

                if (control.HasFile)
                {
                    string fileExtention = System.IO.Path.GetExtension(control.FileName).ToLower();
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

                    String path = System.Web.Hosting.HostingEnvironment.MapPath(virtualDownloadPath);
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    string extention = Path.GetExtension(control.PostedFile.FileName);
                    string filePath = path + fileName + extention;
                    control.PostedFile.SaveAs(filePath);


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
                    string e = "File type must have extentions containing types:";

                    for (int i = 0; i < allowedExtentions.Length; i++)
                    {
                        e += allowedExtentions[i] + ", ";
                    }

                    throw new ArgumentException(e);
                }
            }

            /// <summary>
            /// Serializes a PDF to a byte array.
            /// </summary>
            /// <param name="control">The FileUpload control.</param>
            /// <param name="fileName">Name of the file.</param>
            /// <returns></returns>
            public static byte[] SerializePDF(FileUpload control, string fileName, string virtualDownloadPath = "~/UploadedForms/")
            {
                return SerializeFile(control, fileName, new string[] { ".pdf" }, virtualDownloadPath);
            }

            /// <summary>
            /// Serializes an image.
            /// </summary>
            /// <param name="control">The FileUpload control.</param>
            /// <param name="fileName">Name of the file.</param>
            /// <returns></returns>
            public static byte[] SerializeImage(FileUpload control, string fileName, string virtualDownloadPath = "~/UploadedForms/")
            {
                string[] extentions = new string[] { ".jpg", ".jpeg", ".png", ".bmp" };

                return SerializeFile(control, fileName, extentions, virtualDownloadPath);
            }

            /// <summary>
            /// Deserializes the file.
            /// </summary>
            /// <param name="response">The HttpResponse that will download the File.</param>
            /// <param name="fileName">Name of the file.</param>
            /// <param name="serializedData">The serialized data.</param>
            /// <param name="fileType">Type of the file.</param>
            public static void DeserializeFile(HttpResponse response, string fileName, byte[] serializedData, string fileType)
            {
                response.Clear();
                response.AddHeader("content-disposition",
                "attachment;filename=" + fileName + fileType);
                response.Charset = "";
                response.ContentType = "application/pdf";

                response.Buffer = true;

                //convert to the pdf
                MemoryStream ms = new MemoryStream(serializedData);
                ms.WriteTo(response.OutputStream);
                response.End();
            }

            /// <summary>
            /// Deserializes byte array data, and downloads it as a PDF file.
            /// </summary>
            /// <param name="response">The HttpResponse that will download the PDF File.</param>
            /// <param name="fileName">Name of the file.</param>
            /// <param name="pdfData">The PDF data in its serialized form.</param>
            public static void DeserializePDF(HttpResponse response, string fileName, byte[] pdfData)
            {
                DeserializeFile(response, fileName, pdfData, ".pdf");
            }



        }

    }
}
