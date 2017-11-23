Option Strict Off
Option Explicit On

Imports System.Windows.Forms

Namespace Scripts

    ''' <summary>
    ''' The ribbon code used for the addin
    ''' </summary>
    ''' <remarks></remarks>
    <Runtime.InteropServices.ComVisible(True)>
    Public Class Ribbon
        Implements Office.IRibbonExtensibility
        Private ribbon As Office.IRibbonUI
        Public Shared ribbonref As Ribbon

        Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
            Return GetResourceText("ServerActions.Ribbon.xml")
        End Function

        Private Shared Function GetResourceText(ByVal resourceName As String) As String
            Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
            Dim resourceNames() As String = asm.GetManifestResourceNames()
            For i As Integer = 0 To resourceNames.Length - 1
                If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                    Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                        If resourceReader IsNot Nothing Then
                            Return resourceReader.ReadToEnd()
                        End If
                    End Using
                End If
            Next
            Return Nothing
        End Function

        Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
            Try
                ErrorHandler.CreateLogRecord()
                Me.ribbon = ribbonUI
                ribbonref = Me

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Function GetButtonImage(control As Office.IRibbonControl) As System.Drawing.Bitmap
            Try
                Select Case control.Id
                    Case "btnPing"
                        Return My.Resources.Resources.cmd
                    Case "btnCloudAll"
                        Return My.Resources.Resources.download
                    Case "btnCreateRdgFile"
                        Return My.Resources.Resources.rdg
                    Case Else
                        Return Nothing

                End Select
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

                Return Nothing
            End Try

        End Function

        Public Function GetItemLabel(control As Office.IRibbonControl, index As Integer) As String
            Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveCell.ActiveCell.ListObject
            Try
                If (tbl Is Nothing) Or index = 0 Then
                    ErrorHandler.CreateLogRecord("EMPTY")
                    Return String.Empty
                End If
                ErrorHandler.CreateLogRecord(tbl.ListColumns(index).Name)
                Return tbl.ListColumns(index).Name

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)
                Return "ERROR"

            Finally
                tbl = Nothing

            End Try

        End Function

        Public Function GetItemCount(control As Office.IRibbonControl) As Integer
            Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveCell.ListObject
            Try
                If (tbl Is Nothing) Then
                    Return 2
                Else
                    Return tbl.ListColumns.Count + 1
                End If

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)
                Return 0

            Finally
                tbl = Nothing

            End Try

        End Function

        Public Function GetLabelText(ByVal control As Office.IRibbonControl) As String
            Try
                Select Case control.Id.ToString
                    Case Is = "tabServerActions"
                        If Application.ProductVersion.Substring(0, 2) = "15" Then
                            Return My.Application.Info.Title.ToUpper()
                        Else
                            Return My.Application.Info.Title
                        End If
                    Case Is = "txtCopyright"
                        Return "© " & My.Application.Info.Copyright.ToString
                    Case Is = "txtDescription"
                        Dim version As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
                        Return My.Application.Info.Title.ToString.Replace("&", "&&") & Space(1) & version
                    Case Is = "txtReleaseDate"
                        Return My.Settings.App_ReleaseDate.ToString("dd-MMM-yyyy hh:mm tt")
                    Case Is = "cboServerName"
                        Return My.Settings.Ping_ServerName
                    Case Is = "cboRdgServer"
                        Return My.Settings.Rdg_ServerName
                    Case Is = "cboPingName"
                        Return My.Settings.Ping_Results
                    Case Is = "cboRdgDescription"
                        Return My.Settings.Rdg_Description
                    Case Is = "cboRdgComment"
                        Return My.Settings.Rdg_Comment
                    Case Is = "cboRdgGroup"
                        Return My.Settings.Rdg_ServerGroup
                    Case Is = "txtFileName"
                        Return My.Settings.Rdg_FileName
                End Select
                Return String.Empty

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)
                Return String.Empty

            End Try

        End Function

        Public Shared Sub InvalidateRibbon()
            Try
                ribbonref.ribbon.Invalidate()

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub ActivateTab()
            Try
                ribbonref.ribbon.ActivateTab("tabServerActions")

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub OnAction(ByVal control As Office.IRibbonControl)
            Try
                Select Case control.Id
                    Case "btnPing"
                        Call Ribbon_Button.AddPingColumn()

                    Case "btnCreateRdgFile"
                        Call Ribbon_Button.CreateRdgFile()

                    Case "btnDownloadNewVersion"
                        Call Ribbon_Button.DownloadNewVersion()

                    Case "btnOpenNewIssue"
                        Call Ribbon_Button.OpenNewIssue()

                    Case "btnOpenReadMe"
                        Call Ribbon_Button.OpenReadMe()

                    Case "btnSettings"
                        Call Ribbon_Button.OpenSettings()

                    Case "btnRefreshCombobox"
                        Call Ribbon_Button.RefreshCombobox()

                    Case "btnRefreshServerList"
                        Call Ribbon_Button.RefreshServerList()

                End Select

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Private Sub OnChange(ByVal control As Office.IRibbonControl, ByRef text As String)
            Try

                Select Case control.Id
                    Case Is = "cboServerName"
                        My.Settings.Ping_ServerName = text
                    Case Is = "cboPingName"
                        My.Settings.Ping_Results = text
                    Case Is = "cboRdgServer"
                        My.Settings.Rdg_ServerName = text
                    Case Is = "cboRdgDescription"
                        My.Settings.Rdg_Description = text
                    Case Is = "txtFileName"
                        My.Settings.Rdg_FileName = text
                End Select


            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Function GetItem(col As Object, key As Object) As Object
            Try
                GetItem = col(key)

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)
                GetItem = Nothing

            End Try

        End Function

        Public Shared Sub ClearSheetContents()
            Dim wb As Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook
            Dim ws As Excel.Worksheet = wb.ActiveSheet
            Try
                ws.Cells.Clear() 'To clear only content (no formats) use;
                ws.Cells.ClearContents() 'To Clear only formats use;
                ws.Cells.ClearFormats() 'To clear only cell comments use
                ws.Cells.ClearComments()
                ws.Cells.ClearNotes()
                ws.Cells.ClearOutline()
                'ws.UsedRange 'resets the last cell reference

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            Finally
                ws = Nothing

            End Try

        End Sub

        Public Shared Sub CreateTableFromRange(Optional ByVal FirstCellName As String = "A1", Optional ByVal TableStyleName As String = "TableStyleMedium15")
            Dim tbl As Excel.ListObject
            Dim rng As Excel.Range
            Try
                rng = Globals.ThisAddIn.Application.ActiveSheet.Range(Globals.ThisAddIn.Application.ActiveSheet.Range(FirstCellName), Globals.ThisAddIn.Application.ActiveSheet.Range(FirstCellName).SpecialCells(Excel.Constants.xlLastCell))
                tbl = Globals.ThisAddIn.Application.ActiveSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rng, , Excel.XlYesNoGuess.xlYes)
                tbl.TableStyle = TableStyleName
                tbl.Name = My.Settings.Rdg_SheetName
                Globals.ThisAddIn.Application.Columns.AutoFit()

                Dim r As Excel.Range
                r = Globals.ThisAddIn.Application.ActiveCell
                Globals.ThisAddIn.Application.ActiveSheet.Range("C2").Select
                With Globals.ThisAddIn.Application.ActiveWindow
                    .FreezePanes = False
                    .ScrollRow = 1
                    .ScrollColumn = 1
                    .FreezePanes = True
                    .ScrollRow = r.Row
                End With
                r.Select()

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            Finally
                tbl = Nothing
                rng = Nothing

            End Try

        End Sub

        Public Shared Function GetPingResult(hostName As String) As String
            Dim ping As Object
            Dim status As Object
            Dim result As String = String.Empty
            Dim dateFormat As String = "dd-MMM-yyyy HH:mm:ss"
            Try
                ping = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("Select * from Win32_PingStatus Where Address = '" & hostName & "'")

                For Each status In ping
                    Select Case status.StatusCode
                        Case 0 : result = "Connected"
                        Case 11001 : result = "Buffer too small"
                        Case 11002 : result = "Destination net unreachable"
                        Case 11003 : result = "Destination host unreachable"
                        Case 11004 : result = "Destination protocol unreachable"
                        Case 11005 : result = "Destination port unreachable"
                        Case 11006 : result = "No resources"
                        Case 11007 : result = "Bad option"
                        Case 11008 : result = "Hardware error"
                        Case 11009 : result = "Packet too big"
                        Case 11010 : result = "Request timed out"
                        Case 11011 : result = "Bad request"
                        Case 11012 : result = "Bad route"
                        Case 11013 : result = "Time-To-Live (TTL) expired transit"
                        Case 11014 : result = "Time-To-Live (TTL) expired reassembly"
                        Case 11015 : result = "Parameter problem"
                        Case 11016 : result = "Source quench"
                        Case 11017 : result = "Option too big"
                        Case 11018 : result = "Bad destination"
                        Case 11032 : result = "Negotiating IPSEC"
                        Case 11050 : result = "General failure"
                        Case Else : result = "Unknown host"
                    End Select
                Next
                Return result & " : " & Format(Now(), dateFormat)

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)
                Return "Error: " & ex.ToString()

            Finally
                ping = Nothing
                status = Nothing

            End Try

        End Function

        Public Shared Sub OpenFile(ByVal fileName As String)
            Dim pStart As New System.Diagnostics.Process
            Try
                If fileName = String.Empty Then Exit Try
                pStart.StartInfo.FileName = fileName
                pStart.Start()

            Catch ex As System.ComponentModel.Win32Exception
                'MessageBox.Show("No application is assicated to this file type." & vbCrLf & vbCrLf & pstrFile, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)
                Exit Try

            Finally
                pStart.Dispose()

            End Try

        End Sub

        Public Shared Sub UpdateBlankCells()
            Dim tbl As Excel.ListObject = Nothing
            Dim cell As Excel.Range = Nothing
            Dim usedRange As Excel.Range = Nothing
            Try
                If ErrorHandler.IsAvailable(True) = False Then
                    Return
                End If
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                cell = Nothing
                Dim cnt As Integer = 0
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                usedRange = tbl.Range
                Dim n As Integer = tbl.ListColumns.Count
                Dim m As Integer = tbl.ListRows.Count
                For i As Integer = 0 To m
                    For j As Integer = 1 To n
                        If usedRange(i + 1, j).Value2 Is Nothing Then
                            cell = tbl.Range.Cells(i + 1, j)
                            cell.Value = "NULL"
                            cnt = cnt + 1
                        End If
                    Next
                Next

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            Finally
                tbl = Nothing
                cell = Nothing

            End Try

        End Sub

        Public Shared Function FirstNotNullCellInColumn(ByVal rng As Excel.Range) As Excel.Range
            Try
                If (rng Is Nothing) Then
                    Return Nothing
                End If
                Dim cell As Excel.Range

                For Each cell In rng
                    If (cell.Value IsNot Nothing) Then
                        If (cell.Value.ToString <> "NULL") Then
                            Return cell
                        End If
                    End If
                Next
                Return Nothing

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)
                Return Nothing

            End Try

        End Function

        Public Shared Sub FormatDateColumns()
            Dim tbl As Excel.ListObject = Nothing
            Dim cell As Excel.Range = Nothing
            Dim dateFormat As String = "dd-MMM-yyyy HH:mm:ss"
            Try
                If ErrorHandler.IsAvailable(True) = False Then
                    Return
                End If
                ErrorHandler.CreateLogRecord()
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                cell = Nothing
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                For Each col As Excel.ListColumn In tbl.ListColumns
                    cell = FirstNotNullCellInColumn(col.DataBodyRange)
                    If ((cell IsNot Nothing)) Then
                        If ErrorHandler.IsDate(cell.Value) Then
                            col.DataBodyRange.NumberFormat = dateFormat
                            col.DataBodyRange.HorizontalAlignment = Excel.Constants.xlCenter
                        End If
                    End If
                Next

            Catch ex As System.Runtime.InteropServices.COMException
                ErrorHandler.DisplayMessage(ex)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If cell IsNot Nothing Then
                    'Marshal.ReleaseComObject(cell)
                End If
            End Try

        End Sub

    End Class

End Namespace