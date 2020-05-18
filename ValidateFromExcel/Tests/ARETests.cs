#region Using
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using ExcelDataReader;
using System.Data;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using ValidateFromExcel.Helper;

#endregion
namespace ValidateFromExcel
{
    [TestClass]
    public class ARETests 
    {
        [TestMethod]
        public void ReadFromExcelSheet()
        {
            int sheetToValidate;
            string firstColumnToSortBy, secondColumnToSortBy;

            // 1. Set out the Column Headers you want to validate - The column should exisit in both Sheets with matching name
            // Order of the headers is not relavent
            #region ColumnHeaders
            string[] columnHeaders = {"Journal Type", "Account Code", "Date", "Period", "Transaction Date",
                "Settlement Currency", "Settlement Amount", "Origial Currency", "Original Amount", "Debit/Credit",
                "Trading Partner/Counterparty (Incl Branc  Analysis Code", "Risk Code Analysis Code",
                "Country of Insured Item  Analysis Code", "Transaction Code"
            };
            #endregion

            // 2. Both files should be in a .CSV format. If not just convert them with File --> Save As..
            #region FileStreams
            FileStream streamActualResult = File.Open(@"C:\Users\umars\Desktop\LoL\ActualResult.csv", FileMode.Open, FileAccess.Read);
            FileStream streamExpectedResult = File.Open(@"C:\Users\umars\Desktop\LoL\ExpectedResult.csv", FileMode.Open, FileAccess.Read);

            #endregion

            #region DataSets
            IExcelDataReader excelDataReader1;
            excelDataReader1 = ExcelReaderFactory.CreateCsvReader(streamActualResult);

            DataSet actualDataSet = excelDataReader1.AsDataSet();

            // 3.1. Specify the sheet and the two columns you want to sort your rows by.
            // Sorting will be applied by first column given and then by the second column
            // Note: For validating the first sheet of the document, pass parameter "0". For the second sheet "1" etc...
            // Note: For Sorting by column "A" pass parameter "Column0", For column "B" pass paramter "Column1" etc..

            sheetToValidate = 0;
            firstColumnToSortBy = "Column0";
            secondColumnToSortBy = "Column1";

            actualDataSet = DataValidation.SortDataSet(actualDataSet, sheetToValidate, firstColumnToSortBy, secondColumnToSortBy);
          
            IExcelDataReader excelDataReader2;
            excelDataReader2 = ExcelReaderFactory.CreateCsvReader(streamExpectedResult);

            DataSet expectedDataSet = excelDataReader2.AsDataSet();

            // 3.2. Specify the sheet and the two columns you want to sort your rows by.
            // Sorting will be applied by first column given and then by the second column
            // Note: For validating the first sheet of the document, pass parameter "0". For the second sheet "1" etc...
            // Note: For Sorting by column "A" pass parameter "Column0", For column "B" pass paramter "Column1" etc..

            sheetToValidate = 0;
            firstColumnToSortBy = "Column0";
            secondColumnToSortBy = "Column1";

            expectedDataSet = DataValidation.SortDataSet(expectedDataSet, sheetToValidate, firstColumnToSortBy, secondColumnToSortBy);
         
            #endregion

            #region LoadLists
            List<string> actualList = DataValidation.ReadDataSheet(actualDataSet, columnHeaders);

            List<string> expectedList = DataValidation.ReadDataSheet(expectedDataSet, columnHeaders);

            #endregion


            DataValidation.PrintValidationResults(expectedList, actualList);
        }

       

       

       

       
    }
}
