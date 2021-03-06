﻿using ExcelDataReader;
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

            // 1. Read the expected, actual files into a dataset
            expectedReportPath = @"C:\Users\umars\Desktop\GL-Validation_TestData\FebExpectedResult_RowsMismatch.csv";
            actualReportPath = @"C:\Users\umars\Desktop\GL-Validation_TestData\FebActual_RowsMismatch.csv";


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

            DataTable expectedTable = expectedDataSet.Tables[0];


            // 2. column headers have been applied to the data set

            string[] expectedTableColumnNames = new string[] { "Journal Codes", "Accounts", "Description", "Original Amount Debit", "Original Amount Credit",
                "Settlement Amount Debit", "Settlement Amount Credit", "Base Amount Debit", "Base Amount Credit",
                "Risk Codes", "Syndicates", "Placement Type", "Original Currency", "Settlement Currency"};

            string[] finalTableColumnNames = new string[] { "Journal Codes", "Accounts", "Date", "Period", "Original Currency", "Original Amount",
            "Settlement Currency", "Settlement Amount", "Credit/Debit", "Syndicate", "Risk Codes", "Placement Type", "CYP", "Z", "Year", "New Account",
            "Base Amount"};

            for (int i = 0; i < expectedTableColumnNames.Length; i++)
            {
                expectedTable.Columns[$"Column{i}"].ColumnName = expectedTableColumnNames[i];
            }

            for (int i = 0; i < finalTableColumnNames.Length; i++)
            {
                actualTable.Columns[$"Column{i}"].ColumnName = finalTableColumnNames[i];
            }

            // 3. Data tables have been tranformed in the same manner

            DataTable transformedExpectedTable = CreateDataTable(finalTableColumnNames);

            DataTable transformedActualTable = CreateDataTable(finalTableColumnNames);

            DataTable aggregatedExpectedTable = null;


            for (int i = 0; i < expectedTable.Rows.Count; i++)
            {
                // 2. Check wether this is a Debit or credit row
                isCredit = ValidateCreditRow(expectedTable, i);

                // 3. Read the rows of this sorted table and store the row in a string [] to be added into the finalTable. will have to add additional
                // Credit/Debit field
                rowToAdd = ReadTableRowFromExpected(expectedTable, i, isCredit, finalTableColumnNames.Length);

                // 4. create a new finalTable and keep adding the rows to it one by one (string [])...
                InsertRowInTable(transformedExpectedTable, rowToAdd);
            }

            for (int i = 0; i < actualTable.Rows.Count; i++)
            {
                rowToAdd = ReadTableRowFromActual(actualTable, i, finalTableColumnNames.Length);

                InsertRowInTable(transformedActualTable, rowToAdd);
            }

            // 4. Sorting has been applied to both tables (take input form the user)

            DataTable sortedExpectedTable = DataValidation.SortDataTable(transformedExpectedTable, transformedExpectedTable.Columns["Journal Codes"], 
                transformedExpectedTable.Columns["Accounts"], transformedExpectedTable.Columns["Credit/Debit"]);

            DataTable sortedActualTable = DataValidation.SortDataTable(transformedActualTable, transformedActualTable.Columns["Journal Codes"],
                transformedActualTable.Columns["Accounts"], transformedActualTable.Columns["Credit/Debit"]);

            /* Some thoughts for the aggregation process
             * 1 - Aggregation will only happen in the Expected file
             * 2 - There should be method that takes the expected table, for the current row tells how many more rows below
             *     have the same journal code and account. p.s. the row variable should be increased by that number...
             * 3 - Need a second function to indicate which one of those rows are credit and which ones debit - put them in a strings[]
             * 4 - Need a third method now to aggregate the credit/debit rows if more than 1
            */

            // 5. Access whether aggregation is required

            bool tableToBeAggregated = isTableToAggregated(sortedExpectedTable);

            if (tableToBeAggregated)
            {
                aggregatedExpectedTable = CreateDataTable(finalTableColumnNames);

                string[] row;
                int numRowsToAggregate;

                for (int currentRow = 0; currentRow < sortedExpectedTable.Rows.Count; currentRow++)
                {
                    // since the table needs to be aggregated, i want to know how many rows do i need to aggregate

                    numRowsToAggregate = NumberOfRowsToAggregate(sortedExpectedTable, currentRow);

                    if (numRowsToAggregate == 0)
                    {
                        row = AggregateRows(sortedExpectedTable, currentRow, numRowsToAggregate);
                        InsertRowInTable(aggregatedExpectedTable, row);
                    }
                    else
                    {
                        row = AggregateRows(sortedExpectedTable, currentRow, numRowsToAggregate);
                        InsertRowInTable(aggregatedExpectedTable, row);
                        currentRow = currentRow + numRowsToAggregate;
                    }
                        
                }
            }

            
            string printExpected = null;
            string printActual = null;

            // 6. Compare the final expected table content with that of Actual sorted table..

            int numberOfMismatches = 0;
            bool currentRowSuccess = true;

            if (!tableToBeAggregated)
            {
                if (sortedExpectedTable.Rows.Count != sortedActualTable.Rows.Count)
                {
                    Console.WriteLine($"ERROR!!! The number of rows differs in the two files.\n " +
                        $" Expected n* of rows in the expected view is: { sortedExpectedTable.Rows.Count}.\n " +
                        $" Actual n* of rows in the actual view is: { sortedActualTable.Rows.Count}.\n ");

                    Console.WriteLine("The final expected view is printed as below. Please sort your Actual result sheet for comparison!!! \n");

                    for (int i = 0; i < sortedExpectedTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < sortedExpectedTable.Columns.Count; j++)
                        {
                            printExpected += string.Concat("|", sortedExpectedTable.Rows[i][j]);
                        }
                        
                        Console.WriteLine($"Row number: {i+1} \t" + printExpected + "\n");
                        printExpected = null;

                    }

                    Assert.Fail("\n The validation cannot continue due to mismatches in the number of rows");
                }


                printExpected = null;

                for (int i = 0; i < sortedActualTable.Rows.Count; i++)
                {
                    Console.WriteLine($"Current Row is number: {i + 1}\n");
                    for (int j = 0; j < sortedActualTable.Columns.Count; j++)
                    {
                        printExpected += string.Concat("|", sortedExpectedTable.Rows[i][j]);
                        printActual += string.Concat("|", sortedActualTable.Rows[i][j]);

                        if (sortedExpectedTable.Rows[i][j].ToString() != sortedActualTable.Rows[i][j].ToString())
                        {
                            Console.WriteLine($"There is a mismatch at column: {sortedExpectedTable.Columns[j].ColumnName}\n " +
                                $"Expected file value: {sortedExpectedTable.Rows[i][j]}\n " +
                                $"Actual file value: {sortedActualTable.Rows[i][j]}");
                            numberOfMismatches++;
                            currentRowSuccess = false;
                        }
                    }
                    if (currentRowSuccess)
                    {
                        Console.WriteLine($"SUCCESS!!! Row {i + 1} in both files has successfully been validated!\n");
                    }
                    else
                    {
                        Console.WriteLine($"ATTENTION!!! Row {i + 1} has {numberOfMismatches} number of mismatches!\n");
                    }
                    Console.WriteLine(printExpected);
                    Console.WriteLine(printActual + "\n\n");
                    printExpected = null;
                    printActual = null;
                    numberOfMismatches = 0;
                    currentRowSuccess = true;
                }
            }
            else
            {
                printExpected = null;

                if (aggregatedExpectedTable.Rows.Count != sortedActualTable.Rows.Count)
                {
                    Console.WriteLine($"ERROR!!! The number of rows differs in the two files.\n " +
                        $" Expected n* of rows in the expected view is: { aggregatedExpectedTable.Rows.Count}.\n " +
                        $" Actual n* of rows in the actual view is: { sortedActualTable.Rows.Count}.\n ");

                    Console.WriteLine("The final expected view is printed as below. Please sort your Actual result sheet for comparison!!! \n");

                    for (int i = 0; i < aggregatedExpectedTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < aggregatedExpectedTable.Columns.Count; j++)
                        {
                            printExpected += string.Concat("|", aggregatedExpectedTable.Rows[i][j]);
                        }                      
                        Console.WriteLine($"Row number: {i + 1} \t" + printExpected + "\n");
                        printExpected = null;
                    }

                    Assert.Fail("\n The validation cannot continue due to mismatches in the number of rows");
                }

                for (int i = 0; i < sortedActualTable.Rows.Count; i++)
                {

                    Console.WriteLine($"Current Row is number: {i + 1}\n");
                    for (int j = 0; j < sortedActualTable.Columns.Count; j++)
                    {
                        //Console.Write(finalTableColumnNames[j] + "\t");
                        printExpected += string.Concat("|", aggregatedExpectedTable.Rows[i][j]);
                        printActual += string.Concat("|", sortedActualTable.Rows[i][j]);

                        if (aggregatedExpectedTable.Rows[i][j].ToString() != sortedActualTable.Rows[i][j].ToString())
                        {
                            Console.WriteLine($"There is a mismatch at column: {aggregatedExpectedTable.Columns[j].ColumnName}\n " +
                                $"Expected file value: {aggregatedExpectedTable.Rows[i][j]}\n " +
                                $"Actual file value: {sortedActualTable.Rows[i][j]}");
                            numberOfMismatches++;
                            currentRowSuccess = false;
                        }
                    }
                    if (currentRowSuccess)
                    {
                        Console.WriteLine($"SUCCESS!!! Row {i + 1} in both files has successfully been validated!\n");
                    }
                    else
                    {
                        Console.WriteLine($"ATTENTION!!! Row {i + 1} has {numberOfMismatches} number of mismatches!\n");
                    }
                    
                    Console.WriteLine(printExpected);
                    Console.WriteLine(printActual + "\n\n");
                    printExpected = null;
                    printActual = null;
                    numberOfMismatches = 0;
                    currentRowSuccess = true;
                }
            }        
           
        }

        private string[] AggregateRows(DataTable sortedExpectedTable, int currentRow, int numRowsToAggregate)
        {
            string[] aggregatedRow = new string[sortedExpectedTable.Columns.Count];
            double sum =0;
            if (numRowsToAggregate == 0)
            {
                for (int columnIndex = 0; columnIndex < sortedExpectedTable.Columns.Count; columnIndex++)
                {
                    aggregatedRow[columnIndex] = sortedExpectedTable.Rows[currentRow][columnIndex].ToString();
                }
            }
            else
            {
                for (int columnIndex = 0; columnIndex < sortedExpectedTable.Columns.Count; columnIndex++)
                {
                    switch (columnIndex)
                    {
                        case 5:
                            for (int rowIndex = currentRow; rowIndex <= currentRow + numRowsToAggregate; rowIndex++)
                            {
                                sum += Convert.ToDouble(sortedExpectedTable.Rows[rowIndex][columnIndex].ToString());
                            }
                            aggregatedRow[columnIndex] = sum.ToString();
                            break;
                        case 7:
                            sum = 0;
                            for (int rowIndex = currentRow; rowIndex <= currentRow + numRowsToAggregate; rowIndex++)
                            {
                                sum += Convert.ToDouble(sortedExpectedTable.Rows[rowIndex][columnIndex].ToString());
                            }
                            aggregatedRow[columnIndex] = sum.ToString();
                            break;
                        case 16:
                            sum = 0;
                            for (int rowIndex = currentRow; rowIndex <= currentRow + numRowsToAggregate; rowIndex++)
                            {
                                sum += Convert.ToDouble(sortedExpectedTable.Rows[rowIndex][columnIndex].ToString());
                            }
                            aggregatedRow[columnIndex] = sum.ToString();
                            break;
                        default:
                            aggregatedRow[columnIndex] = sortedExpectedTable.Rows[currentRow][columnIndex].ToString();
                            break;
                    }
                }
            }            
            return aggregatedRow;
        }

        private int NumberOfRowsToAggregate(DataTable sortedExpectedTable, int currentRow)
        {
            int rowsToAggregate = 0;

            for (int i = currentRow+1; i < sortedExpectedTable.Rows.Count; i++)
            {
                if (sortedExpectedTable.Rows[currentRow]["Journal Codes"].ToString() == sortedExpectedTable.Rows[i]["Journal Codes"].ToString()
                    && sortedExpectedTable.Rows[currentRow]["Accounts"].ToString() == sortedExpectedTable.Rows[i]["Accounts"].ToString()
                    && sortedExpectedTable.Rows[currentRow]["Credit/Debit"].ToString() == sortedExpectedTable.Rows[i]["Credit/Debit"].ToString())
                    rowsToAggregate++;
                else
                    break;
            }
            return rowsToAggregate;

        }

        private bool isTableToAggregated(DataTable finalExpectedTable)
        {
            for (int currentRow = 0; currentRow < finalExpectedTable.Rows.Count; currentRow++)
            {
                for (int j = currentRow + 1; j < finalExpectedTable.Rows.Count; j++)
                {
                    if (finalExpectedTable.Rows[currentRow]["Journal Codes"].ToString() == finalExpectedTable.Rows[j]["Journal Codes"].ToString()
                        && finalExpectedTable.Rows[currentRow]["Accounts"].ToString() == finalExpectedTable.Rows[j]["Accounts"].ToString()
                        && finalExpectedTable.Rows[currentRow]["Credit/Debit"].ToString() == finalExpectedTable.Rows[j]["Credit/Debit"].ToString())
                        return true;
                }                
            }
            return false;           
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

        private void InsertRowInTable(DataTable finalTable, DataRow rowToAdd) => finalTable.Rows.Add(rowToAdd);

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
