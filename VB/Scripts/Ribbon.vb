Option Strict Off
Option Explicit On

Imports System.Windows.Forms
Imports System.Management

Namespace Scripts

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
                ErrorHandler.DisplayMessage(ex)

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
            Dim tbl As Excel.ListObject = Nothing
            Try
                If ErrorHandler.IsValidListObject Then
                    tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                End If
                If (tbl Is Nothing) Or index = 0 Then
                    ErrorHandler.CreateLogRecord("EMPTY")
                    Return String.Empty
                End If
                ErrorHandler.CreateLogRecord(tbl.ListColumns(index).Name)
                Return tbl.ListColumns(index).Name

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return "ERROR"

            Finally
                tbl = Nothing

            End Try

        End Function

        Public Function GetItemCount(control As Office.IRibbonControl) As Integer
            Dim tbl As Excel.ListObject = Nothing
            Try
                If ErrorHandler.IsValidListObject Then
                    tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                End If
                If (tbl Is Nothing) Then
                    Return 2
                Else
                    Return tbl.ListColumns.Count + 1
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return 0

            Finally
                tbl = Nothing

            End Try

        End Function

        Public Function GetLabelText(ByVal control As Office.IRibbonControl) As String
            Try
                Select Case control.Id.ToString
                    Case "tabServerActions"
                        If Application.ProductVersion.Substring(0, 2) = "15" Then
                            Return My.Application.Info.Title.ToUpper()
                        Else
                            Return My.Application.Info.Title
                        End If
                    Case "txtCopyright"
                        Return "© " & My.Application.Info.Copyright.ToString
                    Case "txtDescription"
                        Dim version As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
                        Return My.Application.Info.Title.ToString.Replace("&", "&&") & Space(1) & version
                    Case "txtReleaseDate"
                        Return My.Settings.App_ReleaseDate.ToString("dd-MMM-yyyy hh:mm tt")
                    Case "cboServerName"
                        Return My.Settings.Ping_ServerName
                    Case "cboRdgServer"
                        Return My.Settings.Rdg_ServerName
                    Case "cboPingName"
                        Return My.Settings.Ping_Results
                    Case "cboRdgDescription"
                        Return My.Settings.Rdg_Description
                    Case "cboRdgComment"
                        Return My.Settings.Rdg_Comment
                    Case "cboRdgGroup"
                        Return My.Settings.Rdg_ServerGroup
                    Case "txtFileName"
                        Return My.Settings.Rdg_FileName
                End Select
                Return String.Empty

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
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
                    Case "btnPing" : Ribbon_Button.AddPingColumn()
                    Case "btnCreateRdgFile" : Ribbon_Button.CreateRdgFile()
                    Case "btnOpenNewIssue" : Ribbon_Button.OpenNewIssue()
                    Case "btnOpenReadMe" : Ribbon_Button.OpenReadMe()
                    Case "btnSettings" : Ribbon_Button.OpenSettings()
                    Case "btnRefreshCombobox" : Ribbon_Button.RefreshCombobox()
                    Case "btnRefreshServerList" : Ribbon_Button.RefreshServerList()
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Private Sub OnChange(ByVal control As Office.IRibbonControl, ByRef text As String)
            Try
                Select Case control.Id
                    Case "cboServerName" : My.Settings.Ping_ServerName = text
                    Case "cboPingName" : My.Settings.Ping_Results = text
                    Case "cboRdgServer" : My.Settings.Rdg_ServerName = text
                    Case "cboRdgDescription" : My.Settings.Rdg_Description = text
                    Case "txtFileName" : My.Settings.Rdg_FileName = text
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Function GetItem(col As Object, key As Object) As Object
            Try
                GetItem = col(key)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                GetItem = Nothing

            End Try

        End Function

        Public Shared Sub ClearSheetContents()
            Try
                Dim ws As Excel.Worksheet = CType(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet, Excel.Worksheet)
                ws.Cells.Clear() 'To clear only content (no formats) use;
                ws.Cells.ClearContents() 'To Clear only formats use;
                ws.Cells.ClearFormats() 'To clear only cell comments use
                ws.Cells.ClearComments()
                ws.Cells.ClearNotes()
                ws.Cells.ClearOutline()
                'ws.UsedRange 'resets the last cell reference

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub FormatTableFromRange(Optional ByVal FirstCellName As String = "A1", Optional ByVal TableStyleName As String = "TableStyleMedium15")
            Try
                Dim rng As Excel.Range = Globals.ThisAddIn.Application.ActiveSheet.Range(Globals.ThisAddIn.Application.ActiveSheet.Range(FirstCellName), Globals.ThisAddIn.Application.ActiveSheet.Range(FirstCellName).SpecialCells(Excel.Constants.xlLastCell))
                Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, rng, , Excel.XlYesNoGuess.xlYes)
                tbl.TableStyle = TableStyleName
                tbl.Name = My.Settings.Rdg_SheetName
                Globals.ThisAddIn.Application.Columns.AutoFit()

                Dim r As Excel.Range = Globals.ThisAddIn.Application.ActiveCell
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
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Function GetPingResult(ByVal hostName As String) As String
            Dim resultText As String = String.Empty
            Dim dateFormat As String = "dd-MMM-yyyy HH:mm:ss"

            Try
                Dim query As SelectQuery = New SelectQuery("Win32_PingStatus", "Address='" & hostName & "'")
                Dim searcher As ManagementObjectSearcher = New ManagementObjectSearcher(query)

                For Each result As ManagementObject In searcher.[Get]()
                    If result("StatusCode") IsNot Nothing Then
                        Select Case CType(result("StatusCode"), UInt32)
                            Case 0 : resultText = "Connected"
                            Case 11001 : resultText = "Buffer too small"
                            Case 11002 : resultText = "Destination net unreachable"
                            Case 11003 : resultText = "Destination host unreachable"
                            Case 11004 : resultText = "Destination protocol unreachable"
                            Case 11005 : resultText = "Destination port unreachable"
                            Case 11006 : resultText = "No resources"
                            Case 11007 : resultText = "Bad option"
                            Case 11008 : resultText = "Hardware error"
                            Case 11009 : resultText = "Packet too big"
                            Case 11011 : resultText = "Bad request"
                            Case 11012 : resultText = "Bad route"
                            Case 11013 : resultText = "Time-To-Live (TTL) expired transit"
                            Case 11014 : resultText = "Time-To-Live (TTL) expired reassembly"
                            Case 11015 : resultText = "Parameter problem"
                            Case 11016 : resultText = "Source quench"
                            Case 11017 : resultText = "Option too big"
                            Case 11018 : resultText = "Bad destination"
                            Case 11032 : resultText = "Negotiating IPSEC"
                            Case 11050 : resultText = "General failure"
                            Case Else : resultText = "Unknown host"
                        End Select

                    Else
                        resultText = "Unknown host"

                    End If
                Next

                Return resultText & " : " & DateTime.Now.ToString(dateFormat)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return "Error: " & ex.ToString()

            End Try

        End Function

        Public Shared Sub OpenFile(ByVal fileName As String)
            Dim myProcess As New System.Diagnostics.Process
            Try
                If fileName = String.Empty Then Exit Try
                myProcess.StartInfo.FileName = fileName
                myProcess.Start()

            Catch ex As System.ComponentModel.Win32Exception
                'MessageBox.Show("No application is assicated to this file type." & vbCrLf & vbCrLf & pstrFile, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Exit Try

            Finally
                myProcess.Dispose()

            End Try

        End Sub

        Public Shared Sub UpdateZeroStringToNull()
            Try
                If ErrorHandler.IsAvailable(True) Then
                    Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                    Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveCell.ListObject
                    Dim usedRange As Excel.Range = tbl.Range

                    For r As Integer = 0 To tbl.ListRows.Count
                        For c As Integer = 1 To tbl.ListColumns.Count
                            If usedRange(r + 1, c).Value2 Is Nothing Then
                                Dim cell As Excel.Range = tbl.Range.Cells(r + 1, c)
                                cell.Value = "NULL"
                            End If
                        Next
                    Next

                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow

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
                ErrorHandler.DisplayMessage(ex)
                Return Nothing

            End Try

        End Function

        Public Shared Sub FormatDateColumns()
            Dim dateFormat As String = "dd-MMM-yyyy HH:mm:ss"
            Try
                If ErrorHandler.IsAvailable(True) Then
                    Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                    ErrorHandler.CreateLogRecord()
                    Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveCell.ListObject
                    Dim cell As Excel.Range = Nothing

                    For Each col As Excel.ListColumn In tbl.ListColumns
                        cell = FirstNotNullCellInColumn(col.DataBodyRange)
                        If ((cell IsNot Nothing)) Then
                            If ErrorHandler.IsDate(cell.Value) Then
                                col.DataBodyRange.NumberFormat = dateFormat
                                col.DataBodyRange.HorizontalAlignment = Excel.Constants.xlCenter
                            End If
                        End If
                    Next

                End If

            Catch ex As System.Runtime.InteropServices.COMException
                ErrorHandler.DisplayMessage(ex)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow

            End Try

        End Sub

    End Class

End Namespace