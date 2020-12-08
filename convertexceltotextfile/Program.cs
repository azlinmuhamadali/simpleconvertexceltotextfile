using ExcelDataReader;
//using Microsoft.Extensions.DependencyInjection;
using System;
using System.Data;
using System.IO;
using System.Text;

namespace convertexceltotextfile
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            string processFolderPath = @"C:\Users\user\Documents\Visual Studio 2019\Projects\simpleapps\convertexceltotextfile";
            string inputFilePathName = processFolderPath + @"\senarainama.xlsx";
            string outputFileName = "test.txt";
            string outputFilePathName = processFolderPath + "\\" + outputFileName;
            if (!Directory.Exists(processFolderPath))
            {
                Directory.CreateDirectory(processFolderPath);
            }

            FileStream stream = new FileStream(inputFilePathName, FileMode.Open, FileAccess.Read);

            //1. Reading from a binary Excel file ('97-2003 format; *.xls)
            //IExcelDataReader excelReader = ExcelReaderFactory.CreateReader(stream);

            //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            //3. DataSet - The result of each spreadsheet will be created in the result.Tables
            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };

            var dataSet = excelReader.AsDataSet(conf);
            DataTable dt = dataSet.Tables[0];

            int intTotalRecords = 0;
            string strTajuk = "Senarai Nama";

            string strHCol1 = "";
            string strHCol2 = "";
            string strHCol3 = "";


            //initialize header column position
            int intHCol1Position = 1;
            int intHCol2Position = 14;
            int intHCol3Position = 40;

            StringBuilder sbHeader = new StringBuilder();
            StringBuilder sbBody = new StringBuilder();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //BODY MAKER
                //initialize body column position
                int intBCol1Position = 1;
                int intBCol2Position = 22;
                int intBCol3Position = 142;

                string strBCol1 = "";
                string strBCol2 = "";
                string strBCol3 = "";

                //column1
                strBCol1 = dt.Rows[i]["ID"].ToString();
                intBCol1Position = strBCol1.Length + intBCol1Position - 1;
                //column2
                strBCol2 = dt.Rows[i]["NAMA"].ToString();
                intBCol2Position = (strBCol2.Length + intBCol2Position - 1) - intBCol1Position;
                //column3
                strBCol3 = dt.Rows[i]["NOPHONE"].ToString();
                intBCol3Position = (strBCol3.Length + intBCol3Position - 1) - intBCol2Position - intBCol1Position;

                sbBody.AppendLine(strBCol1.PadLeft(intBCol1Position)
                + strBCol2.PadLeft(intBCol2Position)
                + strBCol3.PadLeft(intBCol3Position));
            }
            intTotalRecords = dt.Rows.Count;

            //HEADER MAKER
            //column1
            strHCol1 = intTotalRecords.ToString();
            intHCol1Position = strHCol1.Length + intHCol1Position - 1;
            //column2
            strHCol2 = strTajuk;
            intHCol2Position = (strHCol2.Length + intHCol2Position - 1) - intHCol1Position;
            //column3
            strHCol3 = DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss");
            intHCol3Position = (strHCol3.Length + intHCol3Position - 1) - intHCol2Position - intHCol1Position;

            sbHeader = sbHeader.AppendLine(strHCol1.PadLeft(intHCol1Position)
                + strHCol2.PadLeft(intHCol2Position)
                + strHCol3.PadLeft(intHCol3Position));

            //write all stringbuilder to a file
            //if (!File.Exists(outputFilePathName))
            //{
            using (var outputstream = File.Create(outputFilePathName))
            {
                using (StreamWriter sw = new StreamWriter(outputstream))
                {
                    sw.Write(sbHeader);
                    sw.Write(sbBody);
                }
            }
            //}
            //else 
            //{
            //    using (var outputstream = File.Create(outputFilePathName))
            //    {
            //        using (StreamWriter sw = new StreamWriter(outputstream))
            //        {
            //            sw.Write(sbHeader);
            //            sw.Write(sbBody);
            //        }
            //    }
            //}

            //6. Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();
        }
    }
}
