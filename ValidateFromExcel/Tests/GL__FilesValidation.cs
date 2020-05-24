using ExcelDataReader;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

using System.Data;
using System.IO;
using ValidateFromExcel.Helper;

namespace ValidateFromExcel.Tests
{
    [TestClass]
    public class GL__FilesValidation
    {
        [TestMethod]
        public void GLComparison()
        {

            string expectedReportPath, actualReportPath;
            string[] rowToAdd;
            bool isCredit;

            // 1. Read the expected file into a dataset
            expectedReportPath = @"C:\Users\umars\Desktop\JanExpectedResult.csv";
            actualReportPath = @"C:\Users\umars\Desktop\JanActual.csv";


            FileStream streamExpectedResult = File.Open(expectedReportPath, FileMode.Open, FileAccess.Read);
            FileStream streamSctualResult = File.Open(actualReportPath, FileMode.Open, FileAccess.Read);

            IExcelDataReader expectedDataReader;
            expectedDataReader = ExcelReaderFactory.CreateCsvReader(streamExpectedResult);

            IExcelDataReader actualDataReader;
            actualDataReader = ExcelReaderFactory.CreateCsvReader(streamSctualResult);

            DataSet expectedDataSet = expectedDataReader.AsDataSet();
            DataSet actualDataSet = actualDataReader.AsDataSet();

            expectedDataReader.Dispose();
            actualDataReader.Dispose();

            DataTable actualTable = actualDataSet.Tables[0];

            DataTable originalTable = expectedDataSet.Tables[0];

            string[] expectedTableColumnNames = new string[] { "Journal Codes", "Accounts", "Description", "Original Amount Debit", "Original Amount Credit",
                "Settlement Amount Debit", "Settlement Amount Credit", "Base Amount Debit", "Base Amount Credit",
                "Risk Codes", "Syndicates", "Placement Type", "Original Currency", "Settlement Currency"};

            string[] finalTableColumnNames = new string[] { "Journal Codes", "Accounts", "Date", "Period", "Original Currency", "Original Amount",
            "Settlement Currency", "Settlement Amount", "Credit/Debit", "Syndicate", "Risk Codes", "Placement Type", "CYP", "Z", "Year", "New Account",
            "Base Amount"};

            for (int i = 0; i < expectedTableColumnNames.Length; i++)
            {
                originalTable.Columns[$"Column{i}"].ColumnName = expectedTableColumnNames[i];
            }

            for (int i = 0; i < finalTableColumnNames.Length; i++)
            {
                actualTable.Columns[$"Column{i}"].ColumnName = finalTableColumnNames[i];
            }

            expectedDataSet = DataValidation.SortDataSet(expectedDataSet, originalTable.Columns["Journal Codes"], originalTable.Columns["Accounts"]);       
            
            actualDataSet = DataValidation.SortDataSet(actualDataSet, actualTable.Columns["Journal Codes"], actualTable.Columns["Accounts"]);

            DataTable sortedActualTable = actualDataSet.Tables["Sorted"];

            DataTable sortedExpectedTable = expectedDataSet.Tables["Sorted"];

            DataTable finalExpectedTable = CreateDataTable(finalTableColumnNames);

            DataTable finalActualTable = CreateDataTable(finalTableColumnNames);


            for (int i = 0; i < sortedExpectedTable.Rows.Count; i++)
            {
                // 2. Check wether this is a Debit or credit row
                isCredit = ValidateCreditRow(sortedExpectedTable, i);

                // 3. Read the rows of this sorted table and store the row in a string [] to be added into the finalTable. will have to add additional
                // Credit/Debit field
                rowToAdd = ReadTableRowFromExpected(sortedExpectedTable, i, isCredit, finalTableColumnNames.Length);

                // 4. create a new finalTable and keep adding the rows to it one by one (string [])...
                InsertRowInTable(finalExpectedTable, rowToAdd);
            }

            for (int i = 0; i < sortedActualTable.Rows.Count; i++)
            {
                rowToAdd = ReadTableRowFromActual(sortedActualTable, i, finalTableColumnNames.Length);

                InsertRowInTable(finalActualTable, rowToAdd);
            }

            /* Some thoughts for the aggregation process
             * 1 - Aggregation will only happen in the Expected file
             * 2 - There should be method that takes the expected table, for the current row tells how many more rows below
             *     have the same journal code and account. p.s. the row variable should be increased by that number...
             * 3 - Need a second function to indicate which one of those rows are credit and which ones debit - put them in a strings[]
             * 4 - Need a third method now to aggregate the credit/debit rows if more than 1
            */

            string printExpected = null;
            string printActual = null;

            // 5. Compare the final expected table content with that of Actual sorted table..

            Assert.AreEqual(finalExpectedTable.Rows.Count, sortedActualTable.Rows.Count, 
                $"The row number in the two files differs. Expected file rows: {finalExpectedTable.Rows.Count}, " +
                $"Actual file rows: {finalActualTable.Rows.Count}");

            for (int i = 0; i < finalExpectedTable.Rows.Count; i++)
            {

                Console.WriteLine($"Current Row is number: {i+1}");
                for (int j = 0; j < finalExpectedTable.Columns.Count; j++)
                {
                    printExpected += string.Concat("|", finalExpectedTable.Rows[i][j]);
                    printActual += string.Concat("|", finalActualTable.Rows[i][j]);

                    Assert.AreEqual(finalExpectedTable.Rows[i][j], finalActualTable.Rows[i][j],
                    $"FAIL:      There has been a mismatch!! Expected Value: {finalExpectedTable.Rows[i][j]}" +
                    $"Actual Value: {finalActualTable.Rows[i][j]} \n" +
                    $"Last mismatch found at Row: {i+1} \n and the sequence of \n" +
                    $"Expected Row:\t {printExpected} \n" +
                    $"Actual Row:\t {printActual} \n\n" +
                    $"Before checking for mismatch in the files, ensure correct sorting algorithm has been applied!!");
                }
                Console.WriteLine($"SUCCESS!!! Row {i + 1} in both files has successfully been validated!");
                Console.WriteLine(printExpected);
                Console.WriteLine(printActual + "\n\n");
                printExpected = null;
                printActual = null;
            }
        }

        private string[] ReadTableRowFromActual(DataTable sortedActualTable, int rowNumber, int length)
        {
            string[] rowContent = new string[length];

            for (int i = 0; i < length; i++)
            {
                switch (i)
                {
                    case 2:
                        rowContent[i] = string.Empty;
                        break;
                    case 3:
                        rowContent[i] = string.Empty;
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
                    default:
                        rowContent[i] = sortedActualTable.Rows[rowNumber][i].ToString();
                        break;
                }
            }
            return rowContent;
        }

        private void InsertRowInTable(DataTable finalTable, string[] rowToAdd) => finalTable.Rows.Add(rowToAdd);

        private string[] ReadTableRowFromExpected(DataTable sortedTable, int rowNumber, bool isCredit, int length)
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

                        rowContent[i] = rowContent[i].Replace(",", "");
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

                        rowContent[i] = rowContent[i].Replace(",", "");
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

                        rowContent[i] = rowContent[i].Replace(",", "");
                        break;
                    default:
                        break;
                }           
            }

            return rowContent;
        }

        private DataTable CreateDataTable(string[] actualTableColumnNames)
        {
            // Create a new DataTable.
            DataTable finalTable = new DataTable("finalTable");
            // Declare variables for DataColumn and DataRow objects.
            DataColumn column;

            for (int i = 0; i < actualTableColumnNames.Length; i++)
            {
                // Create new DataColumn, set DataType,
                // ColumnName and add to DataTable.
                column = new DataColumn();
                column.DataType = System.Type.GetType("System.String");
                column.ColumnName = actualTableColumnNames[i];
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
