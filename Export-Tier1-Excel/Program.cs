using Microsoft.SharePoint.Client;
using System;
using System.Data;
using System.IO;
using System.Security;
using Microsoft.Office.Interop.Excel;
using System.Configuration;

namespace Export_Tier1_Excel
{
    class Program
    {
        public static string siteUrl = ConfigurationManager.AppSettings["SiteURL"];
        public static string username = ConfigurationManager.AppSettings["UserName"];
        public static string password = ConfigurationManager.AppSettings["Password"];

        public static string spListName = ConfigurationManager.AppSettings["ListName"];
        public static string viewName = ConfigurationManager.AppSettings["ViewName"];

        public static string excelName = ConfigurationManager.AppSettings["ExcelName"];
        public static string destinationList = ConfigurationManager.AppSettings["DestinationList"];
        public static string exportTempLocation = ConfigurationManager.AppSettings["ExportTempLocation"];
        public static string exportTempDirectory = ConfigurationManager.AppSettings["ExportTempDirectory"];
        public static string fileName = ConfigurationManager.AppSettings["FileName"];
        static void Main(string[] args)
        {
            try
            {
                // first make directory
                System.IO.Directory.CreateDirectory(exportTempDirectory);

                // create excel columns and pull data from SharePoint
                System.Data.DataTable table = new System.Data.DataTable();
                Program p = new Program();
                table = p.GetDataTableFromListItemCollection();

                #region Export to excel
                p.WriteDataTableToExcel(table, "Status Update", exportTempLocation, "Details");

                using (var ctx = new ClientContext(siteUrl))
                {
                    SecureString passWord = new SecureString();
                    foreach (char c in password.ToCharArray()) passWord.AppendChar(c);
                    ctx.Credentials = new SharePointOnlineCredentials(username, passWord);
                }

                UploadFile();

                // Excel needs to write to local directory for user
                // remove file from directory to cleanup
                System.IO.File.Delete(exportTempLocation);

            /*  Console.WriteLine();
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("List export to excel completed successfully.");
                Console.Read();
            */
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                Console.Read();
            }
        }

        private System.Data.DataTable GetDataTableFromListItemCollection()
        {
            string strWhere = string.Empty;
            string filePath = string.Empty;

            /* fields
            Priority
            Initiative
            Title -- changed to Projects
            Description
            Status
            Due Date  Due_X0020_Date
            Revised Due Date
            Executive Sponsor
            Business Owner
            IT Owner
            Phase
            IT Comments
            PRB Comments
            */

            System.Data.DataTable dtGetReqForm = new System.Data.DataTable();

            // add columns to data table

            dtGetReqForm.Columns.Add("Priority", System.Type.GetType("System.String"));
            dtGetReqForm.Columns.Add("Initiative", System.Type.GetType("System.String"));
            dtGetReqForm.Columns.Add("Project", System.Type.GetType("System.String"));
            dtGetReqForm.Columns.Add("Description", System.Type.GetType("System.String"));
            dtGetReqForm.Columns.Add("Status", System.Type.GetType("System.String"));
            dtGetReqForm.Columns.Add("Due Date", System.Type.GetType("System.String"));
            dtGetReqForm.Columns.Add("Revised Due Date", System.Type.GetType("System.String"));
            dtGetReqForm.Columns.Add("Executive Sponsor", System.Type.GetType("System.String"));
            dtGetReqForm.Columns.Add("Business Owner", System.Type.GetType("System.String"));
            dtGetReqForm.Columns.Add("IT Owner", System.Type.GetType("System.String"));
            dtGetReqForm.Columns.Add("Phase", System.Type.GetType("System.String"));
            dtGetReqForm.Columns.Add("IT Comments", System.Type.GetType("System.String"));
            dtGetReqForm.Columns.Add("PRB Comments", System.Type.GetType("System.String"));
            dtGetReqForm.Columns.Add("PRB Priority", System.Type.GetType("System.String"));

            dtGetReqForm.Columns.Add("Financial Benefit", System.Type.GetType("System.String"));
            dtGetReqForm.Columns.Add("Difficulty", System.Type.GetType("System.String"));


            //get data from SharePoint list
            using (var clientContext = new ClientContext(siteUrl))
            {
                try
                {
                    SecureString passWord = new SecureString();
                    foreach (char c in password.ToCharArray()) passWord.AppendChar(c);
                    clientContext.Credentials = new SharePointOnlineCredentials(username, passWord);
                    //Console.WriteLine("Connecting \"" + siteUrl + "\"");
                    Web Oweb = clientContext.Web;
                    clientContext.Load(Oweb);
                    clientContext.ExecuteQuery();
                    List spList = clientContext.Web.Lists.GetByTitle(spListName);
                    clientContext.Load(spList);
                    clientContext.Load(spList.Views);
                    clientContext.ExecuteQuery();
                    //Console.WriteLine("Getting List: " + spListName);

                    if (spList != null && spList.ItemCount > 0)
                    {
                        View view = spList.Views.GetByTitle(viewName);
                        clientContext.Load(view);
                        clientContext.ExecuteQuery();
                        ViewFieldCollection viewFields = view.ViewFields;
                        clientContext.Load(viewFields);
                        clientContext.ExecuteQuery();

                        CamlQuery query = new CamlQuery();
                        query.ViewXml = "<View><Query>" + view.ViewQuery + "</Query></View>";
                        ListItemCollection listItems = spList.GetItems(query);

                        clientContext.Load(listItems,
                            items => items.Include(
                                item => item["Title"],
                                item => item["Priority"],
                                item => item["Description"],
                                item => item["Initiative"],
                                item => item["Status"],
                                item => item["Due_x0020_Date"],
                                item => item["Revised_x0020_Due_x0020_Date"],
                                item => item["Executive_x0020_Sponsor"],
                                item => item["Business_x0020_Owner"],
                                item => item["IT_x0020_Owner"],
                                item => item["Phase"],
                                item => item["IT_x0020_Comments"],
                                item => item["PRB_x0020_Comments"],
                                item => item["Difficulty"],
                                item => item["Financial_x0020_Benefit"],
                                item => item["PRB_x0020_Priority"]));

                        try
                        {

                            clientContext.ExecuteQuery();
                        }
                        catch(Exception e)
                        {
                            Console.WriteLine("error loading list columns: " + e.InnerException + ". " + e.StackTrace);
                        }

                        if (listItems != null && listItems.Count > 0)
                        {
                            foreach (var item in listItems)
                            {
                                // load look up columns and people explicitly at item level
                                clientContext.Load(item,
                                    it => it["Initiative"],
                                    it => it["Status"],
                                    it => it["Phase"],
                                    it => it["IT_x0020_Owner"]);

                                clientContext.ExecuteQuery();
                                
                                FieldLookupValue initiative = item["Initiative"] as FieldLookupValue;
                                FieldLookupValue status = item["Status"] as FieldLookupValue;
                                FieldLookupValue phase = item["Phase"] as FieldLookupValue;
                                FieldUserValue itOwner = (FieldUserValue)item["IT_x0020_Owner"];


                                DataRow dr = dtGetReqForm.NewRow();

                                dr["Project"] = item["Title"];
                                dr["Priority"] = item["Priority"];
                                dr["Description"] = item["Description"];
                                if (initiative != null)
                                {
                                    dr["Initiative"] = initiative.LookupValue;
                                }
                                else
                                {
                                    dr["Initiative"] = "";
                                }

                                if (status != null)
                                {
                                    dr["Status"] = status.LookupValue;
                                }
                                else
                                {
                                    dr["Status"] = "";
                                }

                                if (item["Due_x0020_Date"] != null)
                                {
                                    DateTime dt = Convert.ToDateTime(item["Due_x0020_Date"].ToString());
                                    dr["Due Date"] = dt.ToString("MM/dd/yyyy");
                                }
                                else
                                    dr["Due Date"] = "";
                                if (item["Revised_x0020_Due_x0020_Date"] != null)
                                {
                                    DateTime dtr = Convert.ToDateTime(item["Revised_x0020_Due_x0020_Date"].ToString());
                                    dr["Revised Due Date"] = dtr.ToString("MM/dd/yyyy");
                                }
                                else
                                    dr["Revised Due Date"] = "";

                                dr["Executive Sponsor"] = item["Executive_x0020_Sponsor"];
                                dr["Business Owner"] = item["Business_x0020_Owner"];

                                if (itOwner != null)
                                {
                                    dr["IT Owner"] = itOwner.LookupValue;
                                }
                                else
                                {
                                    dr["IT Owner"] = "";
                                }

                                if (phase != null)
                                {
                                    dr["Phase"] = phase.LookupValue;
                                }
                                else
                                {
                                    dr["Phase"] = "";
                                }


                                if (item["Financial_x0020_Benefit"] != null)
                                {
                                    dr["Financial Benefit"] = item["Financial_x0020_Benefit"].ToString();
                                }
                                else
                                    dr["Financial Benefit"] = "";
                                
                                if (item["Difficulty"] != null)
                                {
                                    dr["Difficulty"] = item["Difficulty"].ToString();
                                }
                                else
                                    dr["Difficulty"] = "";




                                dr["IT Comments"] = item["IT_x0020_Comments"];
                                dr["PRB Comments"] = item["PRB_x0020_Comments"];
                                dr["PRB Priority"] = item["PRB_x0020_Priority"];

                                dtGetReqForm.Rows.Add(dr);
                            }
/*
                            // add blank row before footer
                            DataRow spacer = dtGetReqForm.NewRow();
                            spacer["Description"] = " ";
                            spacer["IT Comments"] = "";
                            dtGetReqForm.Rows.Add(spacer);

                            //add footers here
                            DataRow footer = dtGetReqForm.NewRow();
                            footer["Description"] = "Status Legend";
                            footer["IT Comments"] = "Phase Legend";
                            dtGetReqForm.Rows.Add(footer);

                            DataRow footer2 = dtGetReqForm.NewRow();
                            footer2["Description"] = "Green – On Target for time, budget, and scope";
                            footer2["IT Comments"] = "Concept";
                            dtGetReqForm.Rows.Add(footer2);

                            DataRow footer3 = dtGetReqForm.NewRow();
                            footer3["Description"] = "Yellow – At risk for time, budget, or scope (specify)";
                            footer3["IT Comments"] = "Planning";
                            dtGetReqForm.Rows.Add(footer3);

                            DataRow footer4 = dtGetReqForm.NewRow();
                            footer4["Description"] = "Red – Needs to be refactored – no way to recover without changing the time, budget, or scope";
                            footer4["IT Comments"] = "Development / Implementation";
                            dtGetReqForm.Rows.Add(footer4);

                            DataRow footer5 = dtGetReqForm.NewRow();
                            footer5["Description"] = "Status Legend";
                            footer5["IT Comments"] = "UAT";
                            dtGetReqForm.Rows.Add(footer5);

                            DataRow footer6 = dtGetReqForm.NewRow();
                            footer6["IT Comments"] = "Deployed";
                            dtGetReqForm.Rows.Add(footer6);

                            DataRow footer7 = dtGetReqForm.NewRow();
                            footer7["IT Comments"] = "Complete (burned in)";
                            dtGetReqForm.Rows.Add(footer7);

                            DataRow footer8 = dtGetReqForm.NewRow();
                            footer8["IT Comments"] = "Hold";
                            dtGetReqForm.Rows.Add(footer8);
                            */
                        }
                    }
                }
                
               catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    if (clientContext != null)
                        clientContext.Dispose();
                }
            }
            return dtGetReqForm;

        }

        public bool WriteDataTableToExcel(System.Data.DataTable dataTable, string worksheetName, string saveAsLocation, string ReporType)
        {
            System.IO.Directory.CreateDirectory(exportTempDirectory);

            //Console.WriteLine("In WriteDataTableToExcel");
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet;
            Microsoft.Office.Interop.Excel.Range excelCellrange;
            // for making Excel visible
            excel.Visible = false;
            excel.DisplayAlerts = false;

            // Creation a new Workbook
            excelworkBook = excel.Workbooks.Add(Type.Missing);

            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
            excelSheet.Name = "Status Update";


            try
            {
                // add phases above columns
                excelSheet.Cells[1, 1] = "Phases: Concept => Planning => Design => Development => UAT => Deployed => Complete => Hold";
                excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[1, 16]].Merge();

                // loop through each row and add values to our sheet
                int rowcount = 2;
                int finalColumn = 1;
                foreach (DataRow datarow in dataTable.Rows)
                {
                    int exclColumn = 1;
                    rowcount += 1;
                    for (int i = 1; i <= dataTable.Columns.Count; i++)
                    {
                        // on the first iteration we add the column headers
                        if (rowcount == 3)
                        {
                            excelSheet.Cells[2, exclColumn] = dataTable.Columns[i - 1].ColumnName;
                        }

                        // if (datarow[i - 1].ToString() != "")
                        excelSheet.Cells[rowcount, exclColumn] = datarow[i - 1].ToString();

                        exclColumn += 1;
                        finalColumn = exclColumn-1;
                    }

                    // highlight row if it has a status
                    if (datarow[13].ToString() != "")
                    {
                        excelCellrange = excelSheet.Range[excelSheet.Cells[rowcount, 1], excelSheet.Cells[rowcount, 16]];
                        excelCellrange.Interior.Color = System.Drawing.ColorTranslator.FromHtml("#EDFFFF");
                        //excelCellrange.EntireRow.Font.Bold = true;
                    }

                    if (datarow[4].ToString() == "Green")
                    {
                        excelCellrange = excelSheet.Range[excelSheet.Cells[rowcount, 5], excelSheet.Cells[rowcount, 5]];
                        FormattingExcelCells(excelCellrange, "#85e085", System.Drawing.Color.Black, false);
                    }
                    else if (datarow[4].ToString() == "Yellow")
                    {
                        excelCellrange = excelSheet.Range[excelSheet.Cells[rowcount, 5], excelSheet.Cells[rowcount, 5]];
                        FormattingExcelCells(excelCellrange, "#ffff80", System.Drawing.Color.Black, false);
                    }
                    else if (datarow[4].ToString() == "Red")
                    {
                        excelCellrange = excelSheet.Range[excelSheet.Cells[rowcount, 5], excelSheet.Cells[rowcount, 5]];
                        FormattingExcelCells(excelCellrange, "#ff6666", System.Drawing.Color.Black, false);
                    }

                    
                } // end of foreach datarow

            

                // now we resize the columns
                excelCellrange = excelSheet.Range[excelSheet.Cells[2, 1], excelSheet.Cells[rowcount, finalColumn]];
                excelCellrange.WrapText = true;

                excelCellrange.Columns[1].ColumnWidth = 8;
                excelCellrange.Columns[2].ColumnWidth = 15;
                excelCellrange.Columns[3].ColumnWidth = 20;
                excelCellrange.Columns[4].ColumnWidth = 38;
                excelCellrange.Columns[5].ColumnWidth = 8;
                excelCellrange.Columns[6].ColumnWidth = 11;
                excelCellrange.Columns[7].ColumnWidth = 13;
                excelCellrange.Columns[8].ColumnWidth = 15;
                excelCellrange.Columns[9].ColumnWidth = 15;
                excelCellrange.Columns[10].ColumnWidth = 10;
                excelCellrange.Columns[11].ColumnWidth = 11;
                excelCellrange.Columns[12].ColumnWidth = 32;
                excelCellrange.Columns[13].ColumnWidth = 33;
                excelCellrange.Columns[14].ColumnWidth = 0;
                excelCellrange.Columns[15].ColumnWidth = 15;
                excelCellrange.Columns[16].ColumnWidth = 10;

                // change vertical align to top
                excelSheet.Range[excelSheet.Cells[2, 1], excelSheet.Cells[rowcount, finalColumn]].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;

                //excelCellrange = excelSheet.Range[excelSheet.Cells[2, 1], excelSheet.Cells[rowcount, finalColumn]];
                //excelCellrange.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;

                //FormattingExcelCells(excelCellrange, "#ffffff", System.Drawing.Color.Black, false);

                // make column names bold
                excelSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                excelSheet.Cells[1, 1].EntireRow.Font.Size = 13;

                excelSheet.Cells[2, 1].EntireRow.Font.Bold = true;
                excelSheet.Cells[2, 1].EntireRow.RowHeight = 40;
                excelSheet.Cells[2, 1].EntireRow.Font.Size = 13;


                //auto filter on
                //excelCellrange = excelSheet.Range[excelSheet.Cells[2, 1], excelSheet.Cells[rowcount, finalColumn]];
                excelCellrange.Cells.AutoFilter(1, Type.Missing, XlAutoFilterOperator.xlAnd, Type.Missing, true);
                //excelSheet.Cells.AutoFilter(1, Type.Missing, XlAutoFilterOperator.xlAnd, Type.Missing, true);

                DateTime today = DateTime.Today;
                string addToName = today.ToString("yyyyMMdd") + ".xlsx";
                exportTempLocation = exportTempLocation + addToName;
                fileName = fileName + addToName;

                // set print header and footer and other page setup options
                excelSheet.PageSetup.LeftHeader = "Tier 1 ICG Project Status";
                excelSheet.PageSetup.LeftFooter = "Last Printed: &D &T";
                excelSheet.PageSetup.CenterFooter = fileName;
                excelSheet.PageSetup.RightFooter = "Page &P of &N";
                excelSheet.PageSetup.PaperSize = XlPaperSize.xlPaperLegal;
                excelSheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                excelSheet.PageSetup.Zoom = false;
                excelSheet.PageSetup.FitToPagesWide = 1;
                excelSheet.PageSetup.FitToPagesTall = false;
                excelSheet.PageSetup.PrintTitleRows = "$1:$2";
                //excelSheet.PageSetup.PrintTitleRows = "$2:$2";
                excelSheet.PageSetup.TopMargin = 36;   // 36 points is .5"
                excelSheet.PageSetup.BottomMargin = 36;
                excelSheet.PageSetup.LeftMargin = 36;
                excelSheet.PageSetup.RightMargin = 36;


               


                //now save the workbook and exit Excel
                //Console.WriteLine("file location is: " + exportTempLocation);
                try
                {

                    excelworkBook.SaveAs(exportTempLocation);
                    //Console.WriteLine("Should have just saved Excel file to: " + exportTempLocation);

                    
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Catch in saving excel file: " + ex.Message + ". " + ex.InnerException + ". " + ex.StackTrace);
                }

                
                               
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Catch in Excel file creation: " + ex.InnerException + ". " + ex.StackTrace);
                return false;
            }
            finally
            {
                excelworkBook.Close();
                excel.Quit();
                
                excelSheet = null;
                excelCellrange = null;
                excelworkBook = null;
            }

        }

        public void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbool)
        {
            range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
            if (IsFontbool == true)
            {
                range.Font.Bold = false;
            }
        }

        public static void UploadFile()
        {
            using (var ctx = new ClientContext(siteUrl))
            {
                SecureString passWord = new SecureString();
                foreach (char c in password.ToCharArray()) passWord.AppendChar(c);
                ctx.Credentials = new SharePointOnlineCredentials(username, passWord);

                //UploadFile(ctx);
            
            Web web = ctx.Web;
            var list = ctx.Web.Lists.GetByTitle("Documents");
            ctx.Load(list.RootFolder);
            ctx.ExecuteQuery();
            string destination = string.Format("/{0}/{1}", "itstatus/Shared Documents/Tier 1 Status Weekly Reports", fileName);
            //Console.WriteLine("Destination: " + destination);
            //Console.WriteLine("FileName: " + fileName);
            //Console.WriteLine("exportTempLocation: " + exportTempLocation);

            if (ctx.HasPendingRequest)
                ctx.ExecuteQuery();

            using (FileStream fs = new FileStream(exportTempLocation, FileMode.Open))
            {
                Microsoft.SharePoint.Client.File.SaveBinaryDirect(ctx, destination, fs, true);
            }
            }


            /*
            var fileCreationInfo = new FileCreationInformation
            {
                Content = System.IO.File.ReadAllBytes(uploadFilePath),
                Overwrite = true,
                Url = Path.Combine(uploadFolderUrl, fileName)
            };

            var list = context.Web.Lists.GetByTitle("Shared Documents");
            var uploadFile = list.RootFolder.Files.Add(fileCreationInfo);
             * 
             
            context.Load(uploadFile);
            try
            {
                context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Catch in uploading file: " + ex.InnerException + ". " + ex.StackTrace);
            }
             * * */
        }

    }
}
