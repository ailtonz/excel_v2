using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;   //Microsoft.Office.Interop.Excel;
//using Word = Microsoft.Office.Interop.Word;
//using System.Data;
using System.Data.SqlClient;

namespace excel_v2
{
    class excel_v2
    {
        static void Main(string[] args)
        {
            //Console.WriteLine(Environment.CurrentDirectory);



            Banco.executarInstrucao("INSERT INTO [APP_WEB].[TBL_ARQUIVOS_TRANSITO_ACORDOS]([nm_arquivo]) VALUES ('teste_de_uso')");



            //SqlDataReader rd = Banco.carregarDados("SELECT [nm_arquivo] FROM [DB_SISCOB].[APP_WEB].[TBL_ARQUIVOS_TRANSITO_ACORDOS]");

            //while (rd.Read())
            //{
            //    Console.WriteLine(rd["NM_ARQUIVO"].ToString());
            //}





            //// Create a list of accounts.
            //var bankAccounts = new List<Account>
            // {
            //     new Account {
            //                   ID = 345678,
            //                   Balance = 541.27
            //                 },
            //     new Account {
            //                   ID = 1230221,
            //                   Balance = -127.44
            //                 }
            // };

            //// Display the list in an Excel spreadsheet.
            //DisplayInExcel(bankAccounts);

            ////Create a Word document that contains an icon that links to
            ////the spreadsheet.
            ////CreateIconInWordDoc();

            Console.ReadKey();

        }

        static void DisplayInExcel(IEnumerable<Account> accounts)
        {
            var excelApp = new Excel.Application();
            // Make the object visible.
            excelApp.Visible = false;

            // Create a new, empty workbook and add it to the collection returned 
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a praticular template. 
            // Because no argument is sent in this example, Add creates a new workbook. 
            excelApp.Workbooks.Add();

            // This example uses a single workSheet. 
            Excel._Worksheet workSheet = excelApp.ActiveSheet;

            // Earlier versions of C# require explicit casting.
            //Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            // Establish column headings in cells A1 and B1.
            workSheet.Cells[1, "A"] = "ID Number";
            workSheet.Cells[1, "B"] = "Current Balance";

            var row = 1;
            foreach (var acct in accounts)
            {
                row++;
                workSheet.Cells[row, "A"] = acct.ID;
                workSheet.Cells[row, "B"] = acct.Balance;
            }

            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();

            // Call to AutoFormat in Visual C# 2010. This statement replaces the 
            // two calls to AutoFit.
            workSheet.Range["A1", "B3"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);

            // Put the spreadsheet contents on the clipboard. The Copy method has one
            // optional parameter for specifying a destination. Because no argument  
            // is sent, the destination is the Clipboard.
            //workSheet.Range["A1:B3"].Copy();

            string sPath = Environment.CurrentDirectory;

            workSheet.SaveAs(sPath + "\\tmp3.xlsx");
            excelApp.Workbooks.Close();

        }

        //static void createiconinworddoc()
        //{
        //    var wordapp = new word.application();
        //    wordapp.visible = true;

        //    // the add method has four reference parameters, all of which are 
        //    // optional. visual c# 2010 allows you to omit arguments for them if
        //    // the default values are what you want.
        //    wordapp.documents.add();

        //    // pastespecial has seven reference parameters, all of which are 
        //    // optional. this example uses named arguments to specify values 
        //    // for two of the parameters. although these are reference 
        //    // parameters, you do not need to use the ref keyword, or to create 
        //    // variables to send in as arguments. you can send the values directly.
        //    wordapp.selection.pastespecial(link: true, displayasicon: true);
        //}

    }

    public class Account
    {
        public int ID { get; set; }
        public double Balance { get; set; }
    }
}
