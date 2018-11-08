using ADODB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Management;
using System.Windows.Forms;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ServerActions.Scripts
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        /// <summary>
        /// Used to reference the ribbon object
        /// </summary>
        public static Ribbon ribbonref;

        #region | Task Panes |

        /// <summary>
        /// Settings TaskPane
        /// </summary>
        public Taskpane.Settings mySettings;

        /// <summary>
        /// Settings Custom Task Pane
        /// </summary>
        public Microsoft.Office.Tools.CustomTaskPane myTaskPaneSettings;

        #endregion

        #region | Ribbon Events |

        /// <summary> 
        /// The ribbon
        /// </summary>
        public Ribbon()
        {
        }

        /// <summary> 
        /// Loads the XML markup, either from an XML customization file or from XML markup embedded in the procedure, that customizes the Ribbon user interface.
        /// </summary>
        /// <param name="ribbonID">Represents the XML customization file </param>
        /// <returns>A method that returns a bitmap image for the control id. </returns> 
        /// <remarks></remarks>
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ServerActions.Ribbon.xml");
        }

        /// <summary>
        /// Called by the GetCustomUI method to obtain the contents of the Ribbon XML file.
        /// </summary>
        /// <param name="resourceName">name of  the XML file</param>
        /// <returns>the contents of the XML file</returns>
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        /// <summary> 
        /// loads the ribbon UI and creates a log record
        /// </summary>
        /// <param name="ribbonUI">Represents the IRibbonUI instance that is provided by the Microsoft Office application to the Ribbon extensibility code. </param>
        /// <remarks></remarks>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            try
            {
                this.ribbon = ribbonUI;
                ribbonref = this;
                AssemblyInfo.SetAddRemoveProgramsIcon("ExcelAddin.ico");
                AssemblyInfo.SetAssemblyFolderVersion();

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary> 
        /// Assigns an image to a button on the ribbon in the xml file
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns a bitmap image for the control id. </returns> 
        public System.Drawing.Bitmap GetButtonImage(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "btnPing":
                        return Properties.Resources.cmd;
                    case "btnCreateRdgFile":
                        return Properties.Resources.rdg;
                    default:
                        return null;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return null;
            }
        }

        /// <summary> 
        /// Assigns text to a label on the ribbon from the xml file
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns a string for a label. </returns> 
        public string GetLabelText(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "tabServerActions":
                        if (Application.ProductVersion.Substring(0, 2) == "15") //for Excel 2013
                        {
                            return AssemblyInfo.Title.ToUpper();
                        }
                        else
                        {
                            return AssemblyInfo.Title;
                        }
                    case "txtCopyright":
                        return "© " + AssemblyInfo.Copyright;
                    case "txtDescription":
                        return AssemblyInfo.Title.Replace("&", "&&") + " " + AssemblyInfo.AssemblyVersion;
                    case "txtReleaseDate":
                        DateTime dteCreateDate = Properties.Settings.Default.App_ReleaseDate;
                        return dteCreateDate.ToString("dd-MMM-yyyy hh:mm tt");





                    default:
                        return string.Empty;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return string.Empty;
            }
        }

        /// <summary> 
        /// Assigns the number of items for a combobox or dropdown
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns an integer of total count of items used for a combobox or dropdown </returns> 
        public int GetItemCount(Office.IRibbonControl control)
        {
            Excel.ListObject tbl = null;
            try
            {

                if (ErrorHandler.IsValidListObject())
                {
                    tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
                }

                if (tbl == null)
                {
                    return 1;
                }

                return tbl.ListColumns.Count;
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return 0;
            }
        }

        /// <summary> 
        /// Assigns the values to a combobox or dropdown based on an index
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <param name="index">Represents the index of the combobox or dropdown value </param>
        /// <returns>A method that returns a string per index of a combobox or dropdown </returns> 
        public string GetItemLabel(Office.IRibbonControl control, int index)
        {
            Excel.ListObject tbl = null;
            try
            {

                if (ErrorHandler.IsValidListObject())
                {
                    tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
                }

                if (tbl == null)
                {
                    return String.Empty;
                }

                return tbl.ListColumns[index + 1].Name;
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return string.Empty;
            }
        }

        /// <summary> 
        /// Assigns default values to comboboxes
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns a string for the default value of a combobox </returns> 
        public string GetText(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "cboServerName":
                        return Properties.Settings.Default.Ping_ServerName;
                    case "cboPingName":
                        return Properties.Settings.Default.Ping_Results;
                    case "cboRdgGroup":
                        return Properties.Settings.Default.Rdg_ServerGroup;
                    case "cboRdgServer":
                        return Properties.Settings.Default.Rdg_ServerName;
                    case "cboRdgDescription":
                        return Properties.Settings.Default.Rdg_Description;
                    case "cboRdgComment":
                        return Properties.Settings.Default.Rdg_Comment;
                    case "txtFileName":
                        return Properties.Settings.Default.Rdg_FileName;
                    default:
                        return string.Empty;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return string.Empty;
            }
        }

        /// <summary>
        /// Assigns the value to an application setting
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <returns>A method that returns true or false if the control is enabled </returns> 
        public void OnAction(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "btnPing":
                        AddPingColumn();
                        break;
                    case "btnCreateRdgFile":
                        CreateRdgFile();
                        break;
                    case "btnDownloadNewVersion":
                        //DownloadNewVersion();
                        break;
                    case "btnOpenNewIssue":
                        OpenNewIssue();
                        break;
                    case "btnOpenReadMe":
                        OpenReadMe();
                        break;
                    case "btnSettings":
                        OpenSettings();
                        break;
                    case "btnRefreshCombobox":
                        RefreshCombobox();
                        break;
                    case "btnRefreshServerList":
                        RefreshServerList();
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }

        }

        /// <summary> 
        /// Return the updated value from the comboxbox
        /// </summary>
        /// <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility. </param>
        /// <param name="text">Represents the text from the combobox value </param>
        public void OnChange(Office.IRibbonControl control, string text)
        {
            try
            {
                switch (control.Id)
                {
                    case "cboServerName":
                        Properties.Settings.Default.Ping_ServerName = text;
                        break;
                    case "cboPingName":
                        Properties.Settings.Default.Ping_Results = text;
                        break;
                    case "cboRdgServer":
                        Properties.Settings.Default.Rdg_ServerName = text;
                        break;
                    case "cboRdgDescription":
                        Properties.Settings.Default.Rdg_Description = text;
                        break;
                    case "cboRdgComment":
                        Properties.Settings.Default.Rdg_Comment = text;
                        break;
                    case "cboRdgGroup":
                        Properties.Settings.Default.Rdg_ServerGroup = text;
                        break;
                    case "txtFileName":
                        Properties.Settings.Default.Rdg_FileName = text;
                        break;
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
            finally
            {
                Properties.Settings.Default.Save();
                ribbon.InvalidateControl(control.Id);
            }
        }

        #endregion

        #region | Ribbon Buttons |

        public void AddPingColumn()
        {
            try
            {
                Excel.ListColumn colResults;

                if (ErrorHandler.IsValidListObject())
                {
                    Excel.ListObject tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;

                    if (ErrorHandler.IsValidListColumn(tbl, Properties.Settings.Default.Ping_Results))
                    {
                        colResults = tbl.ListColumns[Properties.Settings.Default.Ping_Results];
                    }
                    else
                    {
                        colResults = tbl.ListColumns.Add();
                        colResults.Name = Properties.Settings.Default.Ping_Results;
                    }
                    
                    if (ErrorHandler.IsValidListColumn(tbl, Properties.Settings.Default.Ping_ServerName))
                    {
                        Excel.ListColumn colServer = tbl.ListColumns[Properties.Settings.Default.Ping_ServerName];
                        for (int r = 1; r <= tbl.ListRows.Count; r++)
                        {
                            if (colServer.DataBodyRange.Rows[r].EntireRow.Hidden == false)
                            {
                                Excel.Range cellResult = colResults.DataBodyRange.Cells[1].Offset(r - 1, 0);
                                Excel.Range cellServer = colServer.DataBodyRange.Cells[1].Offset(r - 1, 0);
                                cellResult.Value = Ribbon.GetPingResult(cellServer.Value.ToString());
                            }
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public void CreateRdgFile()
        {
            try
            {
                string quote = ((char)34).ToString();
                string previousGroup = string.Empty;
                string currentGroup = string.Empty;
                if (ErrorHandler.IsValidListObject())
                {
                    var script = new StringBuilder()
                        .AppendLine("<?xml version=" + quote + "1.0" + quote + " encoding=" + quote + "UTF-8" + quote + "?>")
                        .AppendLine("<RDCMan programVersion=" + quote + "2.7" + quote + " schemaVersion=" + quote + "3" + quote + ">")
                        .AppendLine("<file>")
                        .AppendLine("<credentialsProfiles />")
                        .AppendLine("<properties>")
                        .AppendLine("<expanded>True</expanded>")
                        .AppendLine("<name>" + Scripts.AssemblyInfo.Title + "</name>")
                        .AppendLine("</properties>");

                    Excel.ListObject tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
                    Excel.ListColumn colServer = tbl.ListColumns[Properties.Settings.Default.Rdg_ServerName];
                    Excel.ListColumn colDesc = tbl.ListColumns[Properties.Settings.Default.Rdg_Description];
                    Excel.ListColumn colComment = tbl.ListColumns[Properties.Settings.Default.Rdg_Comment];
                    Excel.ListColumn colServerGroup = tbl.ListColumns[Properties.Settings.Default.Rdg_ServerGroup];

                    for (int r = 1; r <= tbl.ListRows.Count; r++)
                    {
                        if (colServer.DataBodyRange.Rows[r].EntireRow.Hidden == false)
                        {
                            Excel.Range cellServer = colServer.DataBodyRange.Cells[1].Offset(r - 1, 0);
                            Excel.Range cellDesc = colDesc.DataBodyRange.Cells[1].Offset(r - 1, 0);
                            Excel.Range cellComment = colComment.DataBodyRange.Cells[1].Offset(r - 1, 0);
                            Excel.Range cellServerGroup = colServerGroup.DataBodyRange.Cells[1].Offset(r - 1, 0);
                            currentGroup = cellServerGroup.Value;

                            if (currentGroup != previousGroup)
                            {
                                script.AppendLine("<group>");
                                script.AppendLine("<properties>");
                                script.AppendLine("<expanded>True</expanded>");
                                script.AppendLine("<name>" + currentGroup + "</name>");
                                script.AppendLine("</properties>");
                            }

                            script.AppendLine("<server>");
                            script.AppendLine("<properties>");
                            script.AppendLine("<name>" + cellServer.Value + "</name>");
                            script.AppendLine("<displayName>" + cellServer.Value + " (" + cellDesc.Value + ")</displayName>");
                            script.AppendLine("<comment>" & cellComment.Value & "</comment>");
                            script.AppendLine("</properties>");
                            script.AppendLine("</server>");

                            if (currentGroup != colServerGroup.DataBodyRange.Cells[1].Offset(r, 0))
                            {
                                script.AppendLine("</group>");
                            }

                        }
                        previousGroup = currentGroup;
                    }

                    script.AppendLine("</file>");
                    script.AppendLine("<connected />");
                    script.AppendLine("<favorites />");
                    script.AppendLine("<recentlyUsed />");
                    script.AppendLine("</RDCMan>");

                    System.IO.File.WriteAllText(Properties.Settings.Default.Rdg_FileName, script.ToString());

                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public void RefreshServerList()
        {
            ADODB.Connection cn = new ADODB.Connection();
            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Command cmd = new ADODB.Command();
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws;
            string ldapQry = Properties.Settings.Default.Rdg_LdapQry;
            try
            {
                cn.Open("Provider=ADsDSOObject;");
                ldapQry = ldapQry.Replace("[Rdg.LdapPath]", Properties.Settings.Default.Rdg_LdapPath);
                cmd.ActiveConnection = cn;
                object recs;
                rs = cn.Execute(ldapQry, out recs, 0);

                //bool sheetExists;
                //For Each ws In wb.Sheets
                //    If My.Settings.Rdg_SheetName = ws.Name Then
                //        sheetExists = True
                //        ws.Activate()
                //    End If
                //Next ws

                //If sheetExists = False Then
                //    ws = wb.ActiveSheet
                //    Dim answer As Integer
                //    string msg = "The sheet named '" & My.Settings.Rdg_SheetName & "' does not exist."
                //    msg = msg & vbCrLf & "Would you like to use the current sheet?"
                //    answer = MsgBox(msg, vbYesNo + vbQuestion, "Sheet Not Found")
                //  'MessageBox.Show(msg, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                //    If answer = vbYes Then
                //        ws = wb.ActiveSheet
                //        My.Settings.Rdg_SheetName = wb.ActiveSheet.Name
                //    Else
                //        Exit Try
                //    End If
                //Else
                //ws = wb.Worksheets[Properties.Settings.Default.Rdg_SheetName];
                ws = wb.ActiveSheet;
                //End If

                ws.Activate();
                Ribbon.ClearSheetContents();
                for (int i = 0; i <= rs.Fields.Count - 1; i++)
                {
                    ws.Cells[1, i + 1].Value = rs.Fields[i].Name;
                }
                ws.Range[ws.Cells[1, 1], ws.Cells[1, rs.Fields.Count]].Font.Bold = true;
                ws.Range["A2"].CopyFromRecordset(rs);

                Ribbon.FormatTableFromRange();
                Excel.ListObject tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
                Ribbon.UpdateZeroStringToNull(tbl);
                Ribbon.FormatDateColumns();

                //'create server type column from the first 2 characters of the server name
                //'If (My.Settings.Rdg_ServerGroup = "ServerType" Then
                //'    tbl.ListColumns.Add(3).Name = My.Settings.Rdg_ServerGroup
                //'    tbl.ListColumns(My.Settings.Rdg_ServerGroup).DataBodyRange.FormulaR1C1 = "=UPPER(IFERROR(IF(SEARCH(""-"", [@Name]) > 0, LEFT([@Name], 2), """"), ""(NONE)""))"
                //'    Globals.ThisAddIn.Application.Columns.AutoFit()
                //'End If

                Ribbon.ribbonref.InvalidateRibbon(); //'reset dropdown lists
                Ribbon.ActivateTab();

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public void RefreshCombobox()
        {
            try
            {
                if (ErrorHandler.IsValidListObject())
                {
                    ribbon.Invalidate();
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public void OpenSettings()
        {
            try
            {
                if (myTaskPaneSettings != null)
                {
                    if (myTaskPaneSettings.Visible == true)
                    {
                        myTaskPaneSettings.Visible = false;
                    }
                    else
                    {
                        myTaskPaneSettings.Visible = true;
                    }
                }
                else
                {
                    mySettings = new Taskpane.Settings();
                    myTaskPaneSettings = Globals.ThisAddIn.CustomTaskPanes.Add(mySettings, "Settings for " + Scripts.AssemblyInfo.Title);
                    myTaskPaneSettings.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                    myTaskPaneSettings.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                    myTaskPaneSettings.Width = 675;
                    myTaskPaneSettings.Visible = true;
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public void OpenReadMe()
        {
            try
            {
                ErrorHandler.CreateLogRecord();
                System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathReadMe);
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public void OpenNewIssue()
        {
            try
            {
                ErrorHandler.CreateLogRecord();
                System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathNewIssue);
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        public void ShowServerStatus()
        {
            string FullComputerName = "<Name of Remote Computer>";
            ConnectionOptions options = new ConnectionOptions();
            ManagementScope scope = new ManagementScope("\\\\" + FullComputerName + "\\root\\cimv2", options);
            scope.Connect();
            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_TerminalService");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection queryCollection = searcher.Get();
            foreach (ManagementObject queryObj in queryCollection)
            {
                Console.WriteLine("-----------------------------------");
                Console.WriteLine("Win32_TerminalService instance");
                Console.WriteLine("-----------------------------------");
                Console.WriteLine("Started: {0}", queryObj["Started"]);
                Console.WriteLine("State: {0}", queryObj["State"]);
                Console.WriteLine("Status: {0}", queryObj["Status"]);
            }

        }

        #endregion

        #region | Subroutines |

        public static void ActivateTab()
        {
            try
            {
                ribbonref.ribbon.ActivateTab("tabServerActions");
            }

            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);

            }

        }

        /// <summary>
        /// Used to update/reset the ribbon values
        /// </summary>
        public void InvalidateRibbon()
        {
            ribbon.Invalidate();
        }

        public static void ClearSheetContents()
        {
            try
            {
                Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                ws.Cells.Clear(); //To clear only content (no formats) use
                ws.Cells.ClearContents(); //To Clear only formats use
                ws.Cells.ClearFormats(); //To clear only cell comments use
                ws.Cells.ClearComments();
                ws.Cells.ClearNotes();
                ws.Cells.ClearOutline();

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="hostName"></param>
        /// <returns></returns>
        public static string GetPingResult(string hostName)
        {
            string resultText = string.Empty;
            string dateFormat = "dd-MMM-yyyy HH:mm:ss";
            try
            {
                SelectQuery query = new SelectQuery("Win32_PingStatus", "Address='" + hostName + "'");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
                foreach (ManagementObject result in searcher.Get())
                {
                    if (result["StatusCode"] != null)
                    {
                        switch ((UInt32)result["StatusCode"])
                        {
                            case 0: { resultText = "Connected"; break; }
                            case 11001: { resultText = "Buffer too small"; break; }
                            case 11002: { resultText = "Destination net unreachable"; break; }
                            case 11003: { resultText = "Destination host unreachable"; break; }
                            case 11004: { resultText = "Destination protocol unreachable"; break; }
                            case 11005: { resultText = "Destination port unreachable"; break; }
                            case 11006: { resultText = "No resources"; break; }
                            case 11007: { resultText = "Bad option"; break; }
                            case 11008: { resultText = "Hardware error"; break; }
                            case 11009: { resultText = "Packet too big"; break; }
                            case 11011: { resultText = "Bad request"; break; }
                            case 11012: { resultText = "Bad route"; break; }
                            case 11013: { resultText = "Time-To-Live (TTL) expired transit"; break; }
                            case 11014: { resultText = "Time-To-Live (TTL) expired reassembly"; break; }
                            case 11015: { resultText = "Parameter problem"; break; }
                            case 11016: { resultText = "Source quench"; break; }
                            case 11017: { resultText = "Option too big"; break; }
                            case 11018: { resultText = "Bad destination"; break; }
                            case 11032: { resultText = "Negotiating IPSEC"; break; }
                            case 11050: { resultText = "General failure"; break; }
                            default: { resultText = "Unknown host"; break; }
                        }
                    }
                    else
                    {
                        resultText = "Unknown host";
                    }
                }
                return resultText + " : " + DateTime.Now.ToString(dateFormat);
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return "Error: " + ex.ToString();
            }
        }

        /// <summary> 
        /// Finds dates columns with the paste format settings or date specific columns and updates with date format setting
        /// </summary>
        /// <remarks></remarks>
        public static void FormatDateColumns()
        {
            Excel.ListObject tbl = null;
            Excel.Range cell = null;
            string dateFormat = "dd-MMM-yyyy HH:mm:ss";
            try
            {
                if (ErrorHandler.IsAvailable(true) == false)
                {
                    return;
                }
                ErrorHandler.CreateLogRecord();
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject;
                cell = default(Excel.Range);
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                foreach (Excel.ListColumn col in tbl.ListColumns)
                {
                    cell = FirstNotNullCellInColumn(col.DataBodyRange);
                    col.DataBodyRange.NumberFormat = dateFormat;
                    col.DataBodyRange.HorizontalAlignment = Excel.Constants.xlLeft;
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
            finally
            {
                Cursor.Current = System.Windows.Forms.Cursors.Arrow;
                if (tbl != null)
                    Marshal.ReleaseComObject(tbl);
                if (cell != null)
                    Marshal.ReleaseComObject(cell);
            }
        }

        /// <summary>
        /// Convert a range of cells to a table with a specific table format
        /// </summary>
        public static void FormatTableFromRange()
        {
            string FirstCellName = "A1";
            string tableName = AssemblyInfo.Title + " " + DateTime.Now.ToString("yyyy-MM-ddThh:mm:ss:fffzzz");
            string tableStyle = "TableStyleMedium15";
            try
            {
                if (ErrorHandler.IsValidListObject(false) == true)
                {
                    return;
                }
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                ErrorHandler.CreateLogRecord();

                //create the table
                Excel.Range rng = Globals.ThisAddIn.Application.ActiveSheet.Range(Globals.ThisAddIn.Application.ActiveSheet.Range(FirstCellName), Globals.ThisAddIn.Application.ActiveSheet.Range(FirstCellName).SpecialCells(Excel.Constants.xlLastCell));
                rng.Worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rng, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing).Name = tableName;
                Excel.ListObject tbl = rng.Worksheet.ListObjects[tableName];
                tbl.TableStyle = tableStyle;
                tbl.Name = tableName;

                //format the table
                Globals.ThisAddIn.Application.Columns.AutoFit();
                Globals.ThisAddIn.Application.ActiveSheet.Activate();
                Globals.ThisAddIn.Application.ActiveSheet.Application.ActiveWindow.FreezePanes = false;
                Globals.ThisAddIn.Application.ActiveSheet.Application.ActiveWindow.ScrollRow = 1;
                Globals.ThisAddIn.Application.ActiveSheet.Application.ActiveWindow.SplitRow = 1;
                Globals.ThisAddIn.Application.ActiveSheet.Application.ActiveWindow.ScrollColumn = 1;
                Globals.ThisAddIn.Application.ActiveSheet.Application.ActiveWindow.SplitColumn = 1;
                Globals.ThisAddIn.Application.ActiveSheet.Application.ActiveWindow.FreezePanes = true;

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
            finally
            {
                Cursor.Current = System.Windows.Forms.Cursors.Arrow;
            }
        }

        /// <summary> 
        /// Get the first not null cell
        /// </summary>   
        /// <param name="rng">Represents the cell range </param>
        /// <returns>A method that returns a range </returns> 
        /// <remarks> 
        /// TODO: find a way to do this without looping.
        /// NOTE: SpecialCells is unreliable when called from VBA UDFs (Odd ??!)               
        ///</remarks> 
        public static Excel.Range FirstNotNullCellInColumn(Excel.Range rng)
        {
            try
            {
                if ((rng == null))
                {
                    return null;
                }

                foreach (Excel.Range row in rng.Rows)
                {
                    Excel.Range cell = row.Cells[1, 1];
                    if ((cell.Value != null))
                    {
                        string cellValue = cell.Value2.ToString();
                        if (String.Compare(cellValue, "NULL", true) != 0)
                        {
                            return cell;
                        }
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
                return null;
            }
        }

        /// <summary> 
        /// Change zero string cell values to string "NULL"
        /// </summary>
        /// <remarks></remarks>
        public static void UpdateZeroStringToNull(Excel.ListObject tbl)
        {
            try
            {
                if (ErrorHandler.IsAvailable(true) == false)
                {
                    return;
                }

                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
                ErrorHandler.CreateLogRecord();

                Excel.Range cell = default(Excel.Range);
                Excel.Range usedRange = tbl.Range;

                for (int r = 0; r <= tbl.ListRows.Count; r++) 
                {
                    for (int c = 1; c <= tbl.ListColumns.Count; c++)
                    {
                        if (usedRange[r + 1, c].Value2 == null)
                        {
                            cell = tbl.Range.Cells[r + 1, c];
                            cell.Value = "NULL";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }

        }

        #endregion

    }
}
