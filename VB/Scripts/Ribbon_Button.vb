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
            Dim lstCol As Excel.ListColumn
            Dim tbl As Excel.ListObject
            Dim col As Excel.ListColumn
            'Dim qt As String
            Dim a As Object
            Dim c As Object
            'Dim cc As Object
            Dim cnt As Integer
            Dim i As Integer
            Dim colServer As String
            Dim colPing As String
            Dim cellServer As Excel.Range
            Dim cellPing As Excel.Range
            Try

                colServer = My.Settings.Ping_ServerName
                colPing = My.Settings.Ping_Results

                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                If (tbl Is Nothing) Then
                    MessageBox.Show("Please select a table.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Try
                End If

                lstCol = Ribbon.GetItem(tbl.ListColumns, colPing)
                If lstCol Is Nothing Then
                    lstCol = tbl.ListColumns.Add
                    lstCol.Name = colPing
                End If

                For Each col In tbl.ListColumns
                    If col.Name = colServer Then
                        a = col.DataBodyRange.Value2
                        For i = LBound(a) To UBound(a)
                            c = a(i, 1)
                            cellServer = col.DataBodyRange.Cells(1).Offset(i - 1, 0)
                            cellPing = lstCol.DataBodyRange.Cells(1).Offset(i - 1, 0)
                            If col.DataBodyRange.Rows(i).EntireRow.Hidden = False Then
                                cellPing.Value = Ribbon.GetPingResult(cellServer.Value)
                            End If
                            cnt = cnt + 1
                        Next
                    End If
                Next

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            Finally
                lstCol = Nothing
                tbl = Nothing
                col = Nothing
                cellServer = Nothing
                cellPing = Nothing

            End Try

        End Sub

        Public Shared Sub CreateRdgFile()
            Dim lstCol As Excel.ListColumn
            Dim tbl As Excel.ListObject
            Dim col As Excel.ListColumn
            Dim a As Object
            Dim c As Object
            Dim cnt As Integer
            Dim i As Integer
            Dim colServer As String
            Dim colDesc As String
            Dim cellServer As Excel.Range
            Dim cellDesc As Excel.Range
            Dim FileName As String
            Dim script As String
            Dim Q As String
            Try
                FileName = My.Settings.Rdg_FileName
                colServer = My.Settings.Rdg_ServerName
                colDesc = My.Settings.Rdg_Description

                Q = Chr(34)
                script = "<?xml version=" & Q & "1.0" & Q & " encoding=" & Q & "UTF-8" & Q & "?>"
                script += vbCrLf & "<RDCMan programVersion=" & Q & "2.7" & Q & " schemaVersion=" & Q & "3" & Q & ">"
                script += vbCrLf & "<file>"
                script += vbCrLf & "<credentialsProfiles />"
                script += vbCrLf & "<properties>"
                script += vbCrLf & "<expanded>True</expanded>"
                script += vbCrLf & "<name>" & My.Application.Info.Title & "</name>"
                script += vbCrLf & "</properties>"

                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                If (tbl Is Nothing) Then
                    MessageBox.Show("Please select a table.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Try
                End If

                lstCol = Ribbon.GetItem(tbl.ListColumns, colDesc)

                For Each col In tbl.ListColumns
                    If col.Name = colServer Then
                        a = col.DataBodyRange.Value2
                        For i = LBound(a) To UBound(a)
                            c = a(i, 1)
                            cellServer = col.DataBodyRange.Cells(1).Offset(i - 1, 0)
                            cellDesc = lstCol.DataBodyRange.Cells(1).Offset(i - 1, 0)
                            If col.DataBodyRange.Rows(i).EntireRow.Hidden = False Then
                                script += vbCrLf & "<server>"
                                script += vbCrLf & "<properties>"
                                script += vbCrLf & "<displayName>" & cellServer.Value & " (" & cellDesc.Value & ")</displayName>"
                                script += vbCrLf & "<name>" & cellServer.Value & "</name>"
                                script += vbCrLf & "</properties>"
                                script += vbCrLf & "</server>"
                            End If
                            cnt = cnt + 1
                        Next
                    End If
                Next
                script += vbCrLf & "</file>"
                script += vbCrLf & "<connected />"
                script += vbCrLf & "<favorites />"
                script += vbCrLf & "<recentlyUsed />"
                script += vbCrLf & "</RDCMan>"

                System.IO.File.WriteAllText(FileName, script)

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            Finally
                lstCol = Nothing
                tbl = Nothing
                col = Nothing
                cellServer = Nothing
                cellDesc = Nothing

            End Try

        End Sub

        Public Shared Sub RefreshCombobox()
            Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveCell.ListObject
            Try
                If (tbl Is Nothing) Then
                    MessageBox.Show("Please select a table.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Try
                End If
                Ribbon.InvalidateRibbon()

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            Finally
                tbl = Nothing

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
                Call ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub OpenNewIssue()
            Try
                Call Ribbon.OpenFile(My.Settings.App_PathNewIssue)

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub OpenReadMe()
            Try
                Call Ribbon.OpenFile(My.Settings.App_PathReadMe)

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub RefreshServerList()

            Dim cn As ADODB.Connection
            Dim rs As ADODB.Recordset
            Dim cmd As ADODB.Command
            Dim wb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
            Dim ws As Excel.Worksheet
            Dim tbl As Excel.ListObject
            Dim iCols As Integer = 0
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
                    ws = wb.ActiveSheet
                    Dim answer As Integer
                    msg = "The sheet named '" & My.Settings.Rdg_SheetName & "' does not exist."
                    msg = msg & vbCrLf & "Would you like to use the current sheet?"
                    answer = MsgBox(msg, vbYesNo + vbQuestion, "Sheet Not Found")
                    'MessageBox.Show(msg, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                    If answer = vbYes Then
                        ws = wb.ActiveSheet
                        My.Settings.Rdg_SheetName = wb.ActiveSheet.Name
                    Else
                        Exit Try
                    End If
                Else
                    ws = Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets(My.Settings.Rdg_SheetName)
                End If

                Globals.ThisAddIn.Application.Sheets(My.Settings.Rdg_SheetName).Activate
                Call Ribbon.ClearSheetContents()
                For iCols = 0 To rs.Fields.Count - 1
                    ws.Cells(1, iCols + 1).Value = rs.Fields(iCols).Name
                Next
                ws.Range(ws.Cells(1, 1), ws.Cells(1, rs.Fields.Count)).Font.Bold = True
                ws.Range("A2").CopyFromRecordset(rs)

                Ribbon.CreateTableFromRange()
                Ribbon.UpdateBlankCells()
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
                Call ErrorHandler.DisplayMessage(ex)

            Finally
                wb = Nothing
                rs = Nothing
                cn = Nothing
                ws = Nothing
                tbl = Nothing

            End Try

        End Sub

        Public Shared Sub DownloadNewVersion()

        End Sub

    End Class

End Namespace