using ExcelDataReader;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ValidateFromExcel.Helper;

namespace ValidateFromExcel.Tests
{
    [TestClass]
    public class GL__FilesValidation
    {
        [TestMethod]
        public void GLComparison()
        {

            string expectedReportPath;
            string[] rowToAdd;
            bool isCredit;

            // 1. Read the expected file into a dataset
            expectedReportPath = @"C:\Users\umars\Desktop\JanExpectedResult.csv";

            FileStream streamExpectedResult = File.Open(expectedReportPath, FileMode.Open, FileAccess.Read);

            IExcelDataReader expectedDataReader;
            expectedDataReader = ExcelReaderFactory.CreateCsvReader(streamExpectedResult);

            DataSet expectedDataSet = expectedDataReader.AsDataSet();

            expectedDataReader.Dispose();

            DataTable originalTable = expectedDataSet.Tables[0];

            string[] expectedTableColumnNames = new string[] { "Journal Codes", "Accounts", "Description", "Original Amount Debit", "Original Amount Credit",
                "Settlement Amount Debit", "Settlement Amount Credit", "Base Amount Debit", "Base Amount Credit",
                "Risk Codes", "Syndicates", "Placement Type", "Original Currency", "Settlement Currency"};

            for (int i = 0; i < expectedTableColumnNames.Length; i++)
            {
                originalTable.Columns[$"Column{i}"].ColumnName = expectedTableColumnNames[i];
            }

            expectedDataSet = DataValidation.SortDataSet(expectedDataSet, originalTable.Columns["Journal Codes"], originalTable.Columns["Accounts"]);                    

            DataTable sortedTable = expectedDataSet.Tables["Sorted"];

            string[] finalExpectedTableColumnNames = new string[] { "Journal Code", "Accounts", "Date", "Period", "Original Currency", "Original Amount",
            "Settlement Currency", "Settlement Amount", "Credit/Debit", "Syndicate", "Risk Codes", "Placement Type", "CYP", "Z", "Year", "New Account",
            "Base Amount"};

            DataTable finalTable = CreateDataTable(finalExpectedTableColumnNames);

            expectedDataSet.Tables.Add("finalTable");


            for (int i = 0; i < sortedTable.Rows.Count; i++)
            {
                // 2. Check wether this is a Debit or credit row
                isCredit = ValidateCreditRow(sortedTable, i);

                // 3. Read the rows of this sorted table and store the row in a string [] to be added into the final finalTable. will have to add additioanl 
                // Credit/Debit field
                rowToAdd = ReadTableRow(sortedTable, i, isCredit, finalExpectedTableColumnNames.Length);

                // 4. create a new finalTable and keep adding the rows to it one by one (string [])...
                InsertRowInTable(finalTable, rowToAdd);
            }

            string print = null;

            // 5. Compare the final expected table content with that of Actual sorted table..

            for (int i = 0; i < finalTable.Rows.Count; i++)
            {
                for (int j = 0; j < finalTable.Columns.Count; j++)
                {
                    print += string.Concat("|", finalTable.Rows[i][j]);
                }

                Console.WriteLine(print + "\n\n");
                print = null;
            }
        }

        private void InsertRowInTable(DataTable finalTable, string[] rowToAdd)
        {
            finalTable.Rows.Add(rowToAdd);
        }

        private string[] ReadTableRow(DataTable sortedTable, int rowNumber, bool isCredit, int length)
        {
            string[] rowContent = new string[length];
            string lastChar;

            for (int i = 0; i < length; i++)
            {

                switch(i)
                {
                    case 0:
                        rowContent[i] = sortedTable.Rows[rowNumber][i].ToString();
                        break;
                    case 1:
                        rowContent[i] = sortedTable.Rows[rowNumber][i].ToString();
                        break;
                    case 2:
                        rowContent[i] = string.Empty;
                        break;
                    case 3:
                        rowContent[i] = string.Empty;
                        break;
                    case 4:
                        rowContent[i] = sortedTable.Rows[rowNumber]["Original Currency"].ToString();
                        break;
                    case 5:
                        rowContent[i] = isCredit
                       ? sortedTable.Rows[rowNumber]["Original Amount Credit"].ToString()
                       : sortedTable.Rows[rowNumber]["Original Amount Debit"].ToString();

                        if (isCredit)
                            rowContent[i] = string.Concat("-" + rowContent[i]);

                        lastChar = rowContent[i].Substring(rowContent[i].Length - 1, 1);
                        if (lastChar == "0")
                            rowContent[i] = rowContent[i].TrimEnd('0');
                        break;
                    case 6:
                        rowContent[i] = sortedTable.Rows[rowNumber]["Settlement Currency"].ToString();
                        break;
                    case 7:
                        rowContent[i] = isCredit
                       ? sortedTable.Rows[rowNumber]["Settlement Amount Credit"].ToString()
                       : sortedTable.Rows[rowNumber]["Settlement Amount Debit"].ToString();

                        if (isCredit)
                            rowContent[i] = string.Concat("-" + rowContent[i]);

                        lastChar = rowContent[i].Substring(rowContent[i].Length - 1, 1);
                        if (lastChar == "0")
                            rowContent[i] = rowContent[i].TrimEnd('0');
                        break;
                    case 8:
                        rowContent[i] = isCredit
                            ? "C"
                            : "D";
                        break;
                    case 9:
                        rowContent[i] = sortedTable.Rows[rowNumber]["Syndicates"].ToString();
                        break;
                    case 10:
                        rowContent[i] = sortedTable.Rows[rowNumber]["Risk Codes"].ToString();
                        break;
                    case 11:
                        rowContent[i] = sortedTable.Rows[rowNumber]["Placement Type"].ToString();
                        break;
                    case 12:
                        rowContent[i] = string.Empty;
                        break;
                    case 13:
                        rowContent[i] = string.Empty;
                        break;
                    case 14:
                        rowContent[i] = string.Empty;
                        break;
                    case 15:
                        rowContent[i] = string.Empty;
                        break;
                    case 16:
                        rowContent[i] = isCredit
                       ? sortedTable.Rows[rowNumber]["Base Amount Credit"].ToString()
                       : sortedTable.Rows[rowNumber]["Base Amount Debit"].ToString();

                        if (isCredit)
                            rowContent[i] = string.Concat("-" + rowContent[i]);

                        lastChar = rowContent[i].Substring(rowContent[i].Length - 1, 1);
                        if (lastChar == "0")
                            rowContent[i] = rowContent[i].TrimEnd('0');
                        break;
                    default:
                        break;
                }           
            }

            return rowContent;
        }

        private DataTable CreateDataTable(string[] finalExpectedTableColumnNames)
        {
            // Create a new DataTable.
            DataTable finalTable = new DataTable("finalTable");
            // Declare variables for DataColumn and DataRow objects.
            DataColumn column;

            for (int i = 0; i < finalExpectedTableColumnNames.Length; i++)
            {
                // Create new DataColumn, set DataType,
                // ColumnName and add to DataTable.
                column = new DataColumn();
                column.DataType = System.Type.GetType("System.String");
                column.ColumnName = finalExpectedTableColumnNames[i];
                column.ReadOnly = true;
                column.Unique = false;
                // Add the Column to the DataColumnCollection.
                finalTable.Columns.Add(column);
            }
            return finalTable;
        }

        private bool ValidateCreditRow(DataTable originalTable, int rowNumber)
        {
            bool isCredit;
            string value = originalTable.Rows[rowNumber]["Settlement Amount Credit"].ToString();

            isCredit = string.IsNullOrWhiteSpace(value) ? false : true;

            return isCredit;
        }
    }
}
