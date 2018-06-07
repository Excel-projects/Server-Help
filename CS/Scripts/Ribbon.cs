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
                    return 2;
                }

                return tbl.ListColumns.Count + 1;
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

                if (tbl == null || index == 0)
                {
                    return String.Empty;
                }

                return tbl.ListColumns[index].Name;
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

        /// <summary>
        /// 
        /// </summary>
        public void AddPingColumn()
        {
            Excel.ListColumn lstCol;
            Excel.ListObject tbl;
            Excel.ListColumn col;
            object a;
            object c;
            int cnt;
            int i;
            string colServer;
            string colPing;
            Excel.Range cellServer;
            Excel.Range cellPing;
            try
            {
                if (ErrorHandler.IsValidListObject())
                {
                //lstCol = Ribbon.GetItem(tbl.ListColumns, colPing)
                //If lstCol Is Nothing Then
                //    lstCol = tbl.ListColumns.Add
                //    lstCol.Name = colPing
                //End If

                //For Each col In tbl.ListColumns
                //    If col.Name = colServer Then
                //        a = col.DataBodyRange.Value2
                //        For i = LBound(a) To UBound(a)
                //            c = a(i, 1)
                //            cellServer = col.DataBodyRange.Cells(1).Offset(i - 1, 0)
                //            cellPing = lstCol.DataBodyRange.Cells(1).Offset(i - 1, 0)
                //            If col.DataBodyRange.Rows(i).EntireRow.Hidden = False Then
                //                cellPing.Value = Ribbon.GetPingResult(cellServer.Value)
                //            End If
                //            cnt = cnt + 1
                //        Next
                //    End If
                //Next
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public void CreateRdgFile()
        {
            Excel.ListColumn lstCol; 
            Excel.ListObject tbl;
            Excel.ListColumn col;
            Excel.Range cellServer;
            Excel.Range cellDesc;
            string colServer = Properties.Settings.Default.Rdg_ServerName;
            string colDesc = Properties.Settings.Default.Rdg_Description;
            string FileName = Properties.Settings.Default.Rdg_FileName;
            string script = string.Empty;
            string vbCrLf = string.Empty;
            string Q = ((char)34).ToString();
            object a;
            object c;
            int cnt;
            int i;
            try
            {
                if (ErrorHandler.IsValidListObject())
                {
                    script = "<?xml version=" + Q + "1.0" + Q + " encoding=" + Q + "UTF-8" + Q + "?>";
                    script += vbCrLf + "<RDCMan programVersion=" + Q + "2.7" + Q + " schemaVersion=" + Q + "3" + Q + ">";
                    script += vbCrLf + "<file>";
                    script += vbCrLf + "<credentialsProfiles />";
                    script += vbCrLf + "<properties>";
                    script += vbCrLf + "<expanded>True</expanded>";
                    script += vbCrLf + "<name>" + Scripts.AssemblyInfo.Title + "</name>";
                    script += vbCrLf + "</properties>";

                    //lstCol = Ribbon.GetItem(tbl.ListColumns, colDesc)

                    //For Each col In tbl.ListColumns
                    //    If col.Name = colServer Then
                    //        a = col.DataBodyRange.Value2
                    //        For i = LBound(a) To UBound(a)
                    //            c = a(i, 1)
                    //            cellServer = col.DataBodyRange.Cells(1).Offset(i - 1, 0)
                    //            cellDesc = lstCol.DataBodyRange.Cells(1).Offset(i - 1, 0)
                    //            If col.DataBodyRange.Rows(i).EntireRow.Hidden = False Then
                    //                script += vbCrLf + "<server>"
                    //                script += vbCrLf + "<properties>"
                    //                script += vbCrLf + "<displayName>" & cellServer.Value & " (" & cellDesc.Value & ")</displayName>"
                    //                script += vbCrLf + "<name>" & cellServer.Value & "</name>"
                    //                script += vbCrLf + "</properties>"
                    //                script += vbCrLf + "</server>"
                    //            End If
                    //            cnt = cnt + 1
                    //        Next
                    //    End If
                    //Next
                    script += vbCrLf + "</file>";
                    script += vbCrLf + "<connected />";
                    script += vbCrLf + "<favorites />";
                    script += vbCrLf + "<recentlyUsed />";
                    script += vbCrLf + "</RDCMan>";

                    System.IO.File.WriteAllText(FileName, script);
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }
        
        /// <summary>
        /// 
        /// </summary>
        public void RefreshServerList()
        {
            ADODB.Connection cn = new ADODB.Connection();
            ADODB.Recordset rs = new ADODB.Recordset();
            ADODB.Command cmd = new ADODB.Command();
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws;
            Excel.ListObject tbl;
            int iCols = 0;
            string msg = String.Empty;
            string ldapQry = Properties.Settings.Default.Rdg_LdapQry;
            try
            {
                cn.Open("Provider=ADsDSOObject;");
                ldapQry = ldapQry.Replace("[Rdg.LdapPath]", Properties.Settings.Default.Rdg_LdapPath);
                cmd.CommandText = ldapQry;
                cmd.ActiveConnection = cn;
                object objRecAff = null;
                object objParameters = null;
                rs = cmd.Execute(out objRecAff, ref objParameters, (int)ADODB.CommandTypeEnum.adCmdText);

                bool sheetExists;
                //For Each ws In wb.Sheets
                //    If My.Settings.Rdg_SheetName = ws.Name Then
                //        sheetExists = True
                //        ws.Activate()
                //    End If
                //Next ws

                //If sheetExists = False Then
                //    ws = wb.ActiveSheet
                //    Dim answer As Integer
                //    msg = "The sheet named '" & My.Settings.Rdg_SheetName & "' does not exist."
                //    msg = msg & vbCrLf & "Would you like to use the current sheet?"
                //    answer = MsgBox(msg, vbYesNo + vbQuestion, "Sheet Not Found")
                //    'MessageBox.Show(msg, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                //    If answer = vbYes Then
                //        ws = wb.ActiveSheet
                //        My.Settings.Rdg_SheetName = wb.ActiveSheet.Name
                //    Else
                //        Exit Try
                //    End If
                //Else
                //    ws = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(My.Settings.Rdg_SheetName)
                //End If

                //Globals.ThisAddIn.Application.Sheets(My.Settings.Rdg_SheetName).Activate
                //Ribbon.ClearSheetContents()
                //For iCols = 0 To rs.Fields.Count - 1
                //    ws.Cells(1, iCols + 1).Value = rs.Fields(iCols).Name
                //Next
                //ws.Range(ws.Cells(1, 1), ws.Cells(1, rs.Fields.Count)).Font.Bold = True
                //ws.Range("A2").CopyFromRecordset(rs)

                //Ribbon.CreateTableFromRange()
                //Ribbon.UpdateBlankCells()
                //Ribbon.FormatDateColumns()

                //'create server type column from the first 2 characters of the server name
                //'If My.Settings.Rdg_ServerGroup = "ServerType" Then
                //'    tbl.ListColumns.Add(3).Name = My.Settings.Rdg_ServerGroup
                //'    tbl.ListColumns(My.Settings.Rdg_ServerGroup).DataBodyRange.FormulaR1C1 = "=UPPER(IFERROR(IF(SEARCH(""-"", [@Name]) > 0, LEFT([@Name], 2), """"), ""(NONE)""))"
                //'    Globals.ThisAddIn.Application.Columns.AutoFit()
                //'End If

                //Ribbon.InvalidateRibbon() 'reset dropdown lists
                //Ribbon.ActivateTab()

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public void RefreshCombobox()
        {
            try
            {
                if (ErrorHandler.IsValidListObject())
                {
                    ribbon.Invalidate();
                    ribbon.InvalidateControl("ID1");
                }

            }
            catch (Exception ex)
            {
                ErrorHandler.DisplayMessage(ex);
            }
        }

        /// <summary> 
        /// Opens the settings taskpane
        /// </summary>
        /// <remarks></remarks>
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

        /// <summary> 
        /// Opens an as built help file
        /// </summary>
        /// <remarks></remarks>
        public void OpenReadMe()
        {
            ErrorHandler.CreateLogRecord();
            System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathReadMe);

        }

        /// <summary> 
        /// Opens an as built help file
        /// </summary>
        /// <remarks></remarks>
        public void OpenNewIssue()
        {
            ErrorHandler.CreateLogRecord();
            System.Diagnostics.Process.Start(Properties.Settings.Default.App_PathNewIssue);

        }

        /// <summary>
        /// 
        /// </summary>
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

    }
}
