#region Using
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using ExcelDataReader;
using System.Data;
using System.Collections.Generic;

#endregion
namespace ValidateFromExcel
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void ReadFromExcelSheet()
        {
            #region ColumnHeaders
            string[] columnHeaders = {"Journal Type", "Account Code", "Accounting Period", "Year of Account", "Transaction Date",
                "Settlement Currency", "Settlement Amount", "Origial Currency", "Original Amount", "Debit/Credit",
                "Trading Partner/Counterparty (Incl Branc  Analysis Code", "Risk Code Analysis Code",
                "Country of Insured Item  Analysis Code", "Transaction Code"
            };
            #endregion

            #region FileStreams
            FileStream streamActualResult = File.Open(@"C:\Users\umars\Desktop\LoL\ActualResult.csv", FileMode.Open, FileAccess.Read);
            FileStream streamExpectedResult = File.Open(@"C:\Users\umars\Desktop\LoL\ExpectedResult.csv", FileMode.Open, FileAccess.Read);

            // 2. Convertion from .xlsx to .csv
            #endregion

            #region DataSets
            IExcelDataReader excelDataReader1;
            excelDataReader1 = ExcelReaderFactory.CreateCsvReader(streamActualResult);

            DataSet actualDataSet = excelDataReader1.AsDataSet();

            // 3. Blocks of view+DataTable code need to be a separte functions


            // Create a DataView in order to sort csv file in the right order
            DataView viewActual = actualDataSet.Tables[0].DefaultView;
            // List the column Names to filter
            viewActual.Sort = "Column0, Column1 DESC";
            //Create a DataTable based on the updated view after filtering
            DataTable actualValuesTable = viewActual.ToTable();
            //Give this new Table a Name
            actualValuesTable.TableName = "Sorted";
            //Add the new table to ActualDataSet
            actualDataSet.Tables.Add(actualValuesTable);

            IExcelDataReader excelDataReader2;
            excelDataReader2 = ExcelReaderFactory.CreateCsvReader(streamExpectedResult);

            DataSet expectedDataSet = excelDataReader2.AsDataSet();

            // Create a DataView in order to sort csv file in the right order
            DataView viewExpected = expectedDataSet.Tables[0].DefaultView;
            // List the column Names to filter
            viewExpected.Sort = "Column0, Column1 DESC";
            //Create a DataTable based on the updated view after filtering
            DataTable expectedValuesTable = viewExpected.ToTable();
            //Give this new Table a Name
            expectedValuesTable.TableName = "Sorted";
            //Add the new table to ActualDataSet
            expectedDataSet.Tables.Add(expectedValuesTable);

           
            #endregion

            #region LoadLists
            List<string> actualList = ReadDataSheet(actualDataSet, columnHeaders);

            List<string> expectedList = ReadDataSheet(expectedDataSet, columnHeaders);

            #endregion

            #region DisplayOutput
            for (int i = 0; i < actualList.Count; i++)
            {
                try
                {
                    ConsoleWriteLine($"----- Validating row {i + 1} between Expected and Actual sheet!! -----");
                    Assert.AreEqual(expectedList[i], actualList[i]);
                    ConsoleWriteLine("SUCCESS!!!");
                    ConsoleWriteLine("Expected Result Row Value:       " + expectedList[i]);
                    ConsoleWriteLine("Actual Result Row Value:           " + actualList[i]);

                    ConsoleWriteLine("\n\n");
                }
                catch (AssertFailedException e)
                {
                    ConsoleWriteLine($"FAILURE!!! There is a mismatch at row: {i + 1}");
                    ConsoleWriteLine(e.Message);
                    Assert.Fail();                 
                }
            }
            #endregion
        }

        private List<string> ReadDataSheet(DataSet dataset, string[] columnHeaders)
        {
            var list = new List<string>();

            string currentRowData = null;


            for (int rowNumber = 1; rowNumber < dataset.Tables["Sorted"].Rows.Count; rowNumber++)
            {
                for (int numberOfColumns = 0; numberOfColumns < columnHeaders.Length; numberOfColumns++)
                {
                    currentRowData += string.Concat("|", GetColumnValueFromExcel(dataset, rowNumber, columnHeaders[numberOfColumns]));
                }

                list.Add(currentRowData);

                currentRowData = null;
            }

            return list;
        }

        private string GetColumnValueFromExcel(DataSet result, int row, string columnName)
        {
            string columnValue = "";
            for (int columnNumber = 0; columnNumber < result.Tables["Sorted"].Columns.Count; columnNumber++)
            {
                if (result.Tables["Sorted"].Rows[0][columnNumber].ToString() == columnName)
                    columnValue = result.Tables["Sorted"].Rows[row][columnNumber].ToString();
            }
            // 1. Apply Data Formatting - living data format
            return columnValue;
        }

        private void ConsoleWriteLine(string output) => Console.WriteLine(DateTime.Now.ToString("hh:mm:ss") + " -- " + output);
    }
}
