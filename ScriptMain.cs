/*
   Microsoft SQL Server Integration Services Script Task
   Write scripts using Microsoft Visual C# 2008.
   The ScriptMain is the entry point class of the script.
*/
/*====================================================================================
 * AUTHOR: Frederick Awuah-Gyasi
 * DATE  : 12.29.2013
 * TITLE: SSIS PROJECT
 *  
 *              
 * REFRENCE: 
 *          1.http://code.msdn.microsoft.com/office/Imoprt-Data-from-Excel-to-705ecfcd
 *          2.http://www.codeproject.com/Tips/636719/
 *                                  Import-MS-Excel-data-to-SQL-Server-table-using-Csh
 * ====================================================================================
 */

using System;
using System.Data;
using System.Text;
using System.Collections.Generic;
using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace ST_ce3279385006458085d6d01ff3dc276d.csproj
{
    [System.AddIn.AddIn("ScriptMain", Version = "1.0", Publisher = "", Description = "")]
    public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
    {

        #region VSTA generated code
        enum ScriptResults
        {
            Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
            Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
        };
        #endregion

        /*
		The execution engine calls this method when the task executes.
		To access the object model, use the Dts property. Connections, variables, events,
		and logging features are available as members of the Dts property as shown in the following examples.

		To reference a variable, call Dts.Variables["MyCaseSensitiveVariableName"].Value;
		To post a log entry, call Dts.Log("This is my log text", 999, null);
		To fire an event, call Dts.Events.FireInformation(99, "test", "hit the help message", "", 0, true);

		To use the connections collection use something like the following:
		ConnectionManager cm = Dts.Connections.Add("OLEDB");
		cm.ConnectionString = "Data Source=localhost;Initial Catalog=AdventureWorks;Provider=SQLNCLI10;Integrated Security=SSPI;Auto Translate=False;";

		Before returning from this method, set the value of Dts.TaskResult to indicate success or failure.
		
		To open Help, press F1.
	*/

        public void Main()
        {
            /*
             * EXCEL CONNECTION/OPENING WORKBOOK/READING SHEET
            */
            
            Excel.Application excelApp = new Excel.ApplicationClass();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"C:\\Users\\Freddie\\Documents\\AcctFundMappingTable20121011",//Path to the file
                                                                    Missing.Value,
                                                                    Missing.Value,
                                                                    Missing.Value,
                                                                    Missing.Value,
                                                                    Missing.Value,
                                                                    Missing.Value,
                                                                    Missing.Value,
                                                                    Missing.Value,
                                                                    Missing.Value,
                                                                    Missing.Value,
                                                                    Missing.Value,
                                                                    Missing.Value,
                                                                    Missing.Value, 
                                                                    Missing.Value);

            Excel.Sheets excelSheets = excelWorkbook.Worksheets;
            string currentSheet = "MappingTable";//Sheet Name
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
            Excel.Range range = (Excel.Range)excelWorksheet.UsedRange;

            Array myValues = (Array)range.Cells.Value2;
           
            /* NUMBER OF COLUMS AND ROWS */
            int vertical = myValues.GetLength(0); 
            int horizontal = myValues.GetLength(1);

            /*
             *DICTIONARY TO HOLD COLUMNS(VALUE) PER COLOR(KEY) 
             * EXAMPLE: { 
             *           RED: (COL1,COL2,COL3)
             *           BLUE:(COL1,COL2,COL3)
             *          }
             *          IN THIS CASE THE DICTIONARY IS myColumns
             */
            Dictionary<string, List<string>> myColumns = new Dictionary<string, List<string>>();

            
            /*
              MUST START WITH INDEX = 7 PER THE REQUIREMENT OF THIS PROJECT.
              GET COLUMN NAMES
             */ 
            for (int i = 7; i <= horizontal; i++)
            {
                if ((range.Cells[1, i] as Excel.Range).Value2 != null & (range.Cells[1, i] as Excel.Range).Interior.Color.ToString() != "16777215")
                {
                    // dt.Columns.Add(new DataColumn(myValues.GetValue(1, i).ToString()));

                    string sValue = (range.Cells[1, i] as Excel.Range).Value2.ToString();
                    string cellColor = (range.Cells[1, i] as Excel.Range).Interior.Color.ToString();

                    /*
                     If the color already exists just add the column
                     * Creating a KEY:Value here 
                     * Example : Red: (col1,col2,col3)
                     *           Blue:(col1,col2,col3)
                     */
                    if (myColumns.ContainsKey(cellColor))
                    {
                        myColumns[cellColor].Add(sValue);
                    }
                    else
                    {
                      myColumns.Add(cellColor, new List<string>{sValue});
                    }
                    
                }
            }
            foreach (KeyValuePair<string, List<string>> color in myColumns)
            {

                /*
                 * BUILDING THE COMMA SEPARATED STRING
                 */
                StringBuilder sb = new StringBuilder(); // Used to construct a comma separated string of column names for the select statement
                string separator = String.Empty;

                
                foreach (string columns in color.Value)
                {
                    sb.Append(separator).Append(columns);
                    separator = ",";
                    
                }


                //SQL OPERATIONS
                string sSQLTable = "SSIS_Test";
                string myExcelDataQuery = string.Format("Select {0} from [{1}$]",sb,currentSheet);

                
                   try 
                { 
                    
                    string ExcelConnection = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Freddie\Documents\AcctFundMappingTable20121011.xlsx;Extended Properties=""Excel 8.0;HDR=YES;""";
 
                    // Create Connection to Excel Workbook 
                    using (OleDbConnection connection =
                                 new OleDbConnection(ExcelConnection)) 
                    { 
                        OleDbCommand command = new OleDbCommand (myExcelDataQuery, connection); 
 
                        connection.Open(); 
 
                        // Create DbDataReader to Data Worksheet 
                        using (OleDbDataReader dr = command.ExecuteReader()) 
                        { 
 
                            // SQL Server Connection String 
                            string sqlConnectionString = @"Data Source=.;Initial Catalog=Batch50;Integrated Security=True"; 
 
                            // Bulk Copy to SQL Server 
                            using (SqlBulkCopy bulkCopy = 
                                       new SqlBulkCopy(sqlConnectionString)) 
                            {
                                bulkCopy.DestinationTableName = sSQLTable; 
                                bulkCopy.WriteToServer(dr); 
                               
                            } 
                        } 
                    } 
                }
                   catch 
                   {
                      
                   } 
                //Testing just to display color and related columns
                   MessageBox.Show(color.Key.ToString() + " :" + sb.ToString());
            }
                
            Dts.TaskResult = (int)ScriptResults.Success;
        }
    }
}