using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FillDataFromXL.Logic
{
    public class FileReader
    {
        public DataTable ReadExcelFile(string filePath, int sheetNum)
        {
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
                    {
                        FallbackEncoding = Encoding.GetEncoding(1252),
                        AutodetectSeparators = new[] { ',', ';', '\t' } // Add any additional separator characters
                    }))
                    {
                        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration
                            {
                                UseHeaderRow = true
                            }
                        });

                        // Assuming the Excel file has a single sheet
                        var dataTable = dataSet.Tables[sheetNum];

                        // Remove empty rows
                        for (int i = dataTable.Rows.Count - 1; i >= 0; i--)
                        {
                            var row = dataTable.Rows[i];
                            bool isEmptyRow = true;

                            for (int j = 0; j < dataTable.Columns.Count; j++)
                            {
                                if (!string.IsNullOrWhiteSpace(row[j]?.ToString()))
                                {
                                    isEmptyRow = false;
                                    break;
                                }
                            }

                            if (isEmptyRow)
                            {
                                dataTable.Rows.RemoveAt(i);
                            }
                        }

                        return dataTable;
                    }
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Please Close the Business Mapping File\n" + ex.Message);
                return null;
            }

        }
        public DataSet ReadCompleteFileWithRemoveNoOfRowsFromFile(String inputfile, int NoOflinesToSkipStartFromZero0)
        {
            DataSet inputds = null;
            try
            {

                //   Status.Text = "Process 1: Loding Original File";


                using (var stream = File.Open(inputfile, FileMode.Open, FileAccess.Read))
                {

                    // Auto-detect format, supports:
                    //  - Binary Excel files (2.0-2003 format; *.xls)
                    //  - OpenXml Excel files (2007 format; *.xlsx)
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {

                        // 2. Use the AsDataSet extension method
                        inputds = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {

                            // Gets or sets a value indicating whether to set the DataColumn.DataType 
                            // property in a second pass.
                            UseColumnDataType = true,

                            // Gets or sets a callback to obtain configuration options for a DataTable. 
                            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                            {

                                // Gets or sets a value indicating the prefix of generated column names.
                                EmptyColumnNamePrefix = "Column",

                                // Gets or sets a value indicating whether to use a row from the 
                                UseHeaderRow = true,
                                // data as column names.

                                // Gets or sets a callback to determine which row is the header row. 
                                // Only called when UseHeaderRow = true.
                                ReadHeaderRow = (rowReader) =>
                                {
                                    // F.ex skip the first row and use the 2nd row as column headers:
                                    for (int counter = 0; counter < NoOflinesToSkipStartFromZero0; counter++)
                                        rowReader.Read();
                                },

                                // Gets or sets a callback to determine whether to include the 
                                // current row in the DataTable.
                                FilterRow = (rowReader) =>
                                {
                                    return true;
                                },

                                // Gets or sets a callback to determine whether to include the specific
                                // column in the DataTable. Called once per column after reading the 
                                // headers.
                                FilterColumn = (rowReader, columnIndex) =>
                                {
                                    return true;
                                }
                            }
                        });
                        // The result of each spreadsheet is in result.Tables
                        stream.Dispose();
                        stream.Close();
                    }
                }
                //   Status.Text = "Process 1: Input File Loaded";

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return inputds;

            //DataTable finalFile = new DataTable() { TableName = "MyTableName" };

            //DataTable inptbl = inputds.Tables["Sheet1"];
            //finalFile = inputds.Tables["Sheet1"];
            //  Status.Text = "Process 1: Processing Input";
            //  finalFile.WriteMacro(p + "\\ChequeBookUtility\\Macro.xls");
        }
    }
}
