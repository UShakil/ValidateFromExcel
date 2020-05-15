using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelDataReader.Core;
using ExcelDataReader.Log; 
using System.IO;
using ExcelDataReader;
using System.Data;
using System.Collections;
using System.Collections.Generic;

namespace ValidateFromExcel
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void ReadFromExcelSheet()
        {

            string[] columnHeaders = {"Journal Type", "Account Code", "Accounting Period", "Year of Account", "Transaction Date",
                "Settlement Currency", "Settlement Amount", "Origial Currency", "Original Amount", "Debit/Credit",
                "Trading Partner/Counterparty (Incl Branc  Analysis Code", "Risk Code Analysis Code",
                "Country of Insured Item  Analysis Code", "Transaction Code"
            };

            FileStream streamActualResult = File.Open(@"C:\Users\umars\Desktop\LoL\ActualResult.csv", FileMode.Open, FileAccess.Read);
            FileStream streamExpectedResult = File.Open(@"C:\Users\umars\Desktop\LoL\ExpectedResult.csv", FileMode.Open, FileAccess.Read);

            IExcelDataReader excelDataReader1;
            excelDataReader1 = ExcelReaderFactory.CreateCsvReader(streamActualResult);

            DataSet actualDataSet = excelDataReader1.AsDataSet();

            IExcelDataReader excelDataReader2;
            excelDataReader2 = ExcelReaderFactory.CreateCsvReader(streamExpectedResult);

            DataSet expectedDataSet = excelDataReader1.AsDataSet();

            List<string> actualList = ReadDataSheet(actualDataSet, columnHeaders);

            List<string> expectedList = ReadDataSheet(expectedDataSet, columnHeaders);

            for (int i = 0; i < actualList.Count; i++)
            {
                Assert.AreEqual(expectedList[i], actualList[i], $"Expected Value: {expectedList[i]}, Actual Value: {actualList[i]} ");
                Console.WriteLine("SUCCESS!!!");
                Console.WriteLine("Expected Result Row Value:       " + expectedList[i]);
                Console.WriteLine("Actual Result Row Value:           " + actualList[i]);

                Console.WriteLine("\n\n");
            }

            //Assert.AreEqual(expectedDataSet, actualDataSet);

            //foreach (var item in expectedList)
            //{
            //    Console.WriteLine(item + "   ");
            //}

            //Console.WriteLine("\n\n");

            //foreach (var item in actualList)
            //{
            //    Console.WriteLine(item + "   ");
            //}

        }

        private List<string> ReadDataSheet(DataSet dataset, string[] columnHeaders)
        {

            //dataset.Tables[0].DefaultView.Sort = ""
            var list = new List<string>();

            string currentRowData = null;


            for (int rowNumber = 1; rowNumber < dataset.Tables[0].Rows.Count; rowNumber++)
            {
                for (int numberOfColumns = 0; numberOfColumns < columnHeaders.Length; numberOfColumns++)
                {
                    currentRowData += string.Concat("|", GetColumnValueFromExcel(dataset, rowNumber, columnHeaders[numberOfColumns]));
                    //list.Add(GetColumnValueFromExcel(result, rowNumber, columnHeaders[numberOfColumns]));
                }

                list.Add(currentRowData);

                currentRowData = null;
            }

            return list;
        }

        private string GetColumnValueFromExcel(DataSet result, int row, string columnName)
        {
            string columnValue = "";
            for (int columnNumber = 0; columnNumber < result.Tables[0].Columns.Count; columnNumber++)
            {
                if (result.Tables[0].Rows[0][columnNumber].ToString() == columnName)
                    columnValue = result.Tables[0].Rows[row][columnNumber].ToString();
            }
            return columnValue;
        }
    }
}
