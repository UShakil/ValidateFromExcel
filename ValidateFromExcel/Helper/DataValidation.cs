using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ValidateFromExcel.Helper
{
    public static class DataValidation
    {
        public static void ConsoleWriteLine(string output) => Console.WriteLine(DateTime.Now.ToString("hh:mm:ss") + " -- " + output);

        private static string ApplyFormatting(string columnHeader, string columnValue)
        {
            string formattedString, year, month;

            switch (columnHeader)
            {
                case "Date":
                    if (columnValue.Contains('/'))
                    {
                        formattedString = columnValue.Replace("/", "");
                        formattedString = string.Concat(formattedString, " 00:00:00");
                    }
                    else
                        formattedString = columnValue;
                    break;
                case "Transaction Date":
                    if (columnValue.Contains('/'))
                    {
                        formattedString = columnValue.Replace("/", "");
                        formattedString = string.Concat(formattedString, " 00:00:00");
                    }
                    else
                        formattedString = columnValue;
                    break;
                case "Period":
                    if (columnValue.Contains('/'))
                    {
                        char[] separator = { '/' };
                        string[] date = columnValue.Split(separator, 2);
                        year = date[0].ToString();
                        month = date[1].ToString().TrimStart('0');
                        formattedString = string.Concat(month, year);
                    }
                    else
                        formattedString = columnValue;
                    break;
                case "Original Amount":
                    if (columnValue.Contains('.'))
                    {
                        int index = columnValue.IndexOf('.');
                        formattedString = columnValue.Substring(0, index);
                    }
                    else
                        formattedString = columnValue;
                    break;
                case "Settlement Amount":
                    if (columnValue.Contains('.'))
                    {
                        int index = columnValue.IndexOf('.');
                        formattedString = columnValue.Substring(0, index);
                    }
                    else
                        formattedString = columnValue;
                    break;
                default:
                    formattedString = columnValue;
                    break;
            }

            return formattedString;
        }

        public static  DataSet SortDataSet(DataSet dataSet, int sheetToValidate, string sortColumn1, string sortColoumn2)
        {
            // Create a DataView in order to sort csv file in the right order
            DataView viewActual = dataSet.Tables[sheetToValidate].DefaultView;
            // List the column Names to filter
            viewActual.Sort = $"{sortColumn1}, {sortColoumn2} DESC";
            //Create a DataTable based on the updated view after filtering
            DataTable actualValuesTable = viewActual.ToTable();
            //Give this new Table a Name
            actualValuesTable.TableName = "Sorted";
            //Add the new table to ActualDataSet
            dataSet.Tables.Add(actualValuesTable);

            return dataSet;
        }

        public static List<string> ReadDataSheet(DataSet dataset, string[] columnHeaders)
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

        internal static void PrintValidationResults(List<string> expectedList, List<string> actualList)
        {
            for (int i = 0; i < actualList.Count; i++)
            {
                try
                {
                    DataValidation.ConsoleWriteLine($"----- Validating row {i + 1} between Expected and Actual sheet!! -----");
                    Assert.AreEqual(expectedList[i], actualList[i]);
                    DataValidation.ConsoleWriteLine("SUCCESS!!!");
                    DataValidation.ConsoleWriteLine("Expected Result Row Value:       " + expectedList[i]);
                    DataValidation.ConsoleWriteLine("Actual Result Row Value:           " + actualList[i]);

                    DataValidation.ConsoleWriteLine("\n\n");
                }
                catch (AssertFailedException e)
                {
                    DataValidation.ConsoleWriteLine($"FAILURE!!! There is a mismatch at row: {i + 1}");
                    DataValidation.ConsoleWriteLine(e.Message);
                    Assert.Fail();
                }
            }
        }

        private static string GetColumnValueFromExcel(DataSet result, int row, string columnName)
        {
            string columnValue = "";
            for (int columnNumber = 0; columnNumber < result.Tables["Sorted"].Columns.Count; columnNumber++)
            {
                if (result.Tables["Sorted"].Rows[0][columnNumber].ToString() == columnName)
                    columnValue = result.Tables["Sorted"].Rows[row][columnNumber].ToString();
            }
            return ApplyFormatting(columnName, columnValue);
        }

    }
}
