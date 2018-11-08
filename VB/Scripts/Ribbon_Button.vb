Option Strict Off
Option Explicit On

Imports System.Windows.Forms
Imports System.Data.OleDb
Imports ADODB

Namespace Scripts

    Public Class Ribbon_Button

        Public Shared mySettings As TaskPane.Settings
        Public Shared myTaskPaneSettings As Microsoft.Office.Tools.CustomTaskPane

        Public Shared Sub AddPingColumn()
            Try
                If ErrorHandler.IsValidListObject Then
                    Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveCell.ListObject
                    Dim colResults As Excel.ListColumn = tbl.ListColumns(tbl.Range.Columns.Count)

                    If colResults.Name <> My.Settings.Ping_Results Then
                        colResults = tbl.ListColumns.Add
                        colResults.Name = My.Settings.Ping_Results
                    End If

                    Dim colServer As Excel.ListColumn = tbl.ListColumns(My.Settings.Ping_ServerName)

                    For r = 1 To tbl.ListRows.Count
                        If colServer.DataBodyRange.Rows(r).EntireRow.Hidden = False Then
                            Dim cellServer As Excel.Range = colServer.DataBodyRange.Cells(1).Offset(r - 1, 0)
                            Dim cellPing As Excel.Range = colResults.DataBodyRange.Cells(1).Offset(r - 1, 0)
                            cellPing.Value = Ribbon.GetPingResult(cellServer.Value.ToString)
                        End If
                    Next

                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub CreateRdgFile()
            Dim col As Excel.ListColumn
            Dim quote As String : quote = Chr(34)
            Dim a As Object
            Dim c As Object
            Dim i As Integer
            Try
                If ErrorHandler.IsValidListObject Then
                    Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveCell.ListObject
                    Dim script As String = "<?xml version=" + quote + "1.0" + quote + " encoding=" + quote + "UTF-8" + quote + "?>"
                    script += vbCrLf + "<RDCMan programVersion=" + quote + " 2.7" + quote + " schemaVersion=" + quote + " 3" + quote + ">"
                    script += vbCrLf + "<file>"
                    script += vbCrLf + "<credentialsProfiles/>"
                    script += vbCrLf + "<properties>"
                    script += vbCrLf + "<expanded>True</expanded>"
                    script += vbCrLf + "<name>" + My.Application.Info.Title + "</name>"
                    script += vbCrLf + "</properties>"

                    Dim lstCol As Excel.ListColumn = Ribbon.GetItem(tbl.ListColumns, My.Settings.Rdg_Description)

                    For Each col In tbl.ListColumns
                        If col.Name = My.Settings.Rdg_ServerName Then
                            a = col.DataBodyRange.Value2
                            For i = LBound(a) To UBound(a)
                                c = a(i, 1)
                                Dim cellServer As Excel.Range = col.DataBodyRange.Cells(1).Offset(i - 1, 0)
                                Dim cellDesc As Excel.Range = lstCol.DataBodyRange.Cells(1).Offset(i - 1, 0)
                                If col.DataBodyRange.Rows(i).EntireRow.Hidden = False Then
                                    script += vbCrLf + "<server>"
                                    script += vbCrLf + "<properties>"
                                    script += vbCrLf + "<displayName>" + cellServer.Value.ToString + " (" + cellDesc.Value.ToString + ")</displayName>"
                                    script += vbCrLf + "<name>" + cellServer.Value.ToString + "</name>"
                                    script += vbCrLf + "</properties>"
                                    script += vbCrLf + "</server>"
                                End If
                            Next
                        End If
                    Next

                    script += vbCrLf + "</file>"
                    script += vbCrLf + "<connected/>"
                    script += vbCrLf + "<favorites/>"
                    script += vbCrLf + "<recentlyUsed/>"
                    script += vbCrLf + "</RDCMan>"

                    System.IO.File.WriteAllText(My.Settings.Rdg_FileName, script)

                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                col = Nothing

            End Try

        End Sub

        Public Shared Sub RefreshCombobox()
            Try
                If ErrorHandler.IsValidListObject Then
                    Ribbon.InvalidateRibbon()
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub OpenSettings()
            Try
                If myTaskPaneSettings IsNot Nothing Then
                    If myTaskPaneSettings.Visible = True Then
                        myTaskPaneSettings.Visible = False
                    Else
                        myTaskPaneSettings.Visible = True
                    End If
                Else
                    mySettings = New ServerActions.TaskPane.Settings()
                    myTaskPaneSettings = Globals.ThisAddIn.CustomTaskPanes.Add(mySettings, "Settings for " + My.Application.Info.Title)
                    myTaskPaneSettings.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
                    myTaskPaneSettings.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange
                    myTaskPaneSettings.Width = 675
                    myTaskPaneSettings.Visible = True

                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub OpenNewIssue()
            Try
                Ribbon.OpenFile(My.Settings.App_PathNewIssue)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub OpenReadMe()
            Try
                Ribbon.OpenFile(My.Settings.App_PathReadMe)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub RefreshServerList()
            Dim cn As ADODB.Connection
            Dim rs As ADODB.Recordset
            Dim cmd As ADODB.Command
            Dim wb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
            Dim ws As Excel.Worksheet
            Dim tbl As Excel.ListObject
            Dim i As Integer = 0
            Dim msg As String = String.Empty
            Dim ldapQry As String = My.Settings.Rdg_LdapQry
            Try

                cmd = CreateObject("ADODB.Command")
                cn = CreateObject("ADODB.Connection")
                rs = CreateObject("ADODB.Recordset")
                cn.Open("Provider=ADsDSOObject;")
                ldapQry = Replace(ldapQry, "[Rdg.LdapPath]", My.Settings.Rdg_LdapPath)
                cmd.CommandText = ldapQry
                cmd.ActiveConnection = cn
                rs = cmd.Execute

                Dim sheetExists As Boolean
                For Each ws In wb.Sheets
                    If My.Settings.Rdg_SheetName = ws.Name Then
                        sheetExists = True
                        ws.Activate()
                    End If
                Next ws

                If sheetExists = False Then
                    ws = CType(wb.ActiveSheet, Excel.Worksheet)
                    Dim answer As Integer
                    msg = "The sheet named '" & My.Settings.Rdg_SheetName & "' does not exist."
                    msg = msg & vbCrLf & "Would you like to use the current sheet?"
                    answer = MsgBox(msg, vbYesNo + vbQuestion, "Sheet Not Found")
                    'MessageBox.Show(msg, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                    If answer = vbYes Then
                        ws = CType(wb.ActiveSheet, Excel.Worksheet)
                        My.Settings.Rdg_SheetName = ws.Name
                    Else
                        Exit Try
                    End If
                Else
                    ws = CType(Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(My.Settings.Rdg_SheetName), Excel.Worksheet)
                End If

                ws.Activate()
                Ribbon.ClearSheetContents()
                For i = 0 To rs.Fields.Count - 1
                    ws.Cells(1, i + 1).Value = rs.Fields(i).Name
                Next
                ws.Range(ws.Cells(1, 1), ws.Cells(1, rs.Fields.Count)).Font.Bold = True
                ws.Range("A2").CopyFromRecordset(rs)

                Ribbon.FormatTableFromRange()
                Ribbon.UpdateZeroStringToNull()
                Ribbon.FormatDateColumns()

                'create server type column from the first 2 characters of the server name
                'If My.Settings.Rdg_ServerGroup = "ServerType" Then
                '    tbl.ListColumns.Add(3).Name = My.Settings.Rdg_ServerGroup
                '    tbl.ListColumns(My.Settings.Rdg_ServerGroup).DataBodyRange.FormulaR1C1 = "=UPPER(IFERROR(IF(SEARCH(""-"", [@Name]) > 0, LEFT([@Name], 2), """"), ""(NONE)""))"
                '    Globals.ThisAddIn.Application.Columns.AutoFit()
                'End If

                Ribbon.InvalidateRibbon() 'reset dropdown lists
                Ribbon.ActivateTab()

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                wb = Nothing
                rs = Nothing
                cn = Nothing
                ws = Nothing
                tbl = Nothing

            End Try

        End Sub

    End Class

End Namespace