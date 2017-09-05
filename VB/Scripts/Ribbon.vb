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

        Private mySettings As TaskPane.Settings
        Private myTaskPaneSettings As Microsoft.Office.Tools.CustomTaskPane

#Region "| Ribbon Events |"

        ''' <summary>
        ''' 
        ''' </summary>
        Public Sub New()
        End Sub

        ''' <summary>
        ''' Loads the XML markup, either from an XML customization file or from XML markup embedded in the procedure, that customizes the Ribbon user interface.
        ''' </summary>
        ''' <param name="ribbonID">Represents the XML customization file</param>
        ''' <returns>A method that returns a bitmap image for the control id.</returns>
        ''' <remarks></remarks>
        Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
            Return GetResourceText("ServerActions.Ribbon.xml")
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="resourceName"></param>
        ''' <returns></returns>
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

        ''' <summary>
        ''' Load the ribbon
        ''' </summary>
        ''' <param name="ribbonUI">Represents the IRibbonUI instance that is provided by the Microsoft Office application to the Ribbon extensibility code.</param>
        ''' <remarks></remarks>
        Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
            Try
                Me.ribbon = ribbonUI

            Catch ex As Exception
                Call DisplayMessage(ex)

            End Try

        End Sub

        ''' <summary>
        ''' To assign text to controls on the ribbon from the xml file
        ''' </summary>
        ''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
        ''' <returns>A method that returns a string for a label. </returns>
        ''' <remarks></remarks>
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
                        Dim strVersion As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
                        Return My.Application.Info.Title.ToString.Replace("&", "&&") & Space(1) & strVersion
                    Case Is = "txtReleaseDate"
                        Dim dteCreateDate As DateTime = System.IO.File.GetLastWriteTime(My.Application.Info.DirectoryPath.ToString & "\" & My.Application.Info.AssemblyName.ToString & ".dll") 'get creation date 
                        Return dteCreateDate.ToString("dd-MMM-yyyy hh:mm tt")
                    Case Else
                        Return String.Empty
                End Select

            Catch ex As Exception
                Call DisplayMessage(ex)
                'Console.WriteLine(ex.Message.ToString)
                Return String.Empty

            End Try

        End Function

        Public Sub GetItemLabel(ByVal control As Office.IRibbonControl, index As Integer, ByRef returnedVal As String)
            Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveCell.ActiveCell.ListObject
            Try
                If (tbl Is Nothing) Or index = 0 Then
                    returnedVal = ""
                    Exit Sub
                End If
                returnedVal = tbl.ListColumns(index).Name

            Catch ex As Exception
                Call DisplayMessage(ex)
                returnedVal = "ERROR"

            Finally
                tbl = Nothing

            End Try

        End Sub

        Private Function GetItemCount(ByVal control As Office.IRibbonControl, ByRef Count As Integer) As Integer
            Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveCell.ListObject
            Try
                If (tbl Is Nothing) Then
                    Return 2
                Else
                    Return tbl.ListColumns.Count + 1
                End If

            Catch ex As Exception
                Call DisplayMessage(ex)
                Return 0

            Finally
                tbl = Nothing

            End Try

        End Function

        Private Sub GetText(ByVal control As Office.IRibbonControl, ByRef text As String)
            Try

                Select Case control.Id
                    Case Is = "cboServerName", "cboRdgServer"
                        text = "Server"
                    Case Is = "cboPingName"
                        text = "Ping"
                    Case Is = "cboRdgDescription"
                        text = "Description"
                    Case Is = "txtFileName"
                        text = "H:\ServerListing.rdg"
                End Select

            Catch ex As Exception
                Call DisplayMessage(ex)

            End Try

        End Sub

        Private Sub OnChange(ByVal control As Office.IRibbonControl, ByRef text As String)
            Try

                Select Case control.Id
                    Case Is = "cboServerName"
                        My.Settings.Ping_ServerColumn = text
                    Case Is = "cboPingName"
                        My.Settings.Ping_PingColumn = text
                    Case Is = "cboRdgServer"
                        My.Settings.Rdg_ServerColumn = text
                    Case Is = "cboRdgDescription"
                        My.Settings.Rdg_DescriptionColumn = text
                    Case Is = "txtFileName"
                        My.Settings.Rdg_Filename = text
                End Select


            Catch ex As Exception
                Call DisplayMessage(ex)

            End Try

        End Sub

#End Region

#Region "| Ribbon Buttons |"

        Public Sub AddPingColumn(ByVal control As Office.IRibbonControl)
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

                colServer = My.Settings.Ping_ServerColumn
                colPing = My.Settings.Ping_PingColumn

                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                If (tbl Is Nothing) Then
                    MessageBox.Show("Please select a table.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Try
                End If

                lstCol = GetItem(tbl.ListColumns, colPing)
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
                                cellPing.Value = GetPingResult(cellServer.Value)
                            End If
                            cnt = cnt + 1
                        Next
                    End If
                Next

            Catch ex As Exception
                Call DisplayMessage(ex)

            Finally
                lstCol = Nothing
                tbl = Nothing
                col = Nothing
                cellServer = Nothing
                cellPing = Nothing

            End Try

        End Sub

        Public Sub CreateRdgFile(ByVal control As Office.IRibbonControl)
            Dim lstCol As Excel.ListColumn
            Dim tbl As Excel.ListObject
            Dim col As Excel.ListColumn
            'Dim qt As String
            Dim a As Object
            Dim c As Object
            'Dim cc As Variant
            Dim cnt As Integer
            Dim i As Integer
            Dim colServer As String
            Dim colDesc As String
            Dim cellServer As Excel.Range
            Dim cellDesc As Excel.Range
            Dim FileName As String
            Dim script As String
            'Dim nDestFile As Integer
            'Dim sText As String
            'Dim iRow As Integer
            'Dim iColCount As Integer
            'Dim icol As Integer
            Dim Q As String
            Try

                FileName = My.Settings.Rdg_Filename
                colServer = My.Settings.Rdg_ServerColumn
                colDesc = My.Settings.Rdg_DescriptionColumn

                Q = Chr(34)
                script = "<?xml version=" & Q & "1.0" & Q & " encoding=" & Q & "UTF-8" & Q & "?>"
                script = script & vbCrLf & "<RDCMan programVersion=" & Q & "2.7" & Q & " schemaVersion=" & Q & "3" & Q & ">"
                script = script & vbCrLf & "<file>"
                script = script & vbCrLf & "<credentialsProfiles />"
                script = script & vbCrLf & "<properties>"
                script = script & vbCrLf & "<expanded>True</expanded>"
                script = script & vbCrLf & "<name>" & My.Application.Info.Title & "</name>"
                script = script & vbCrLf & "</properties>"

                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                If (tbl Is Nothing) Then
                    MessageBox.Show("Please select a table.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Try
                End If

                lstCol = GetItem(tbl.ListColumns, colDesc)

                For Each col In tbl.ListColumns
                    If col.Name = colServer Then
                        a = col.DataBodyRange.Value2
                        For i = LBound(a) To UBound(a)
                            c = a(i, 1)
                            cellServer = col.DataBodyRange.Cells(1).Offset(i - 1, 0)
                            cellDesc = lstCol.DataBodyRange.Cells(1).Offset(i - 1, 0)
                            If col.DataBodyRange.Rows(i).EntireRow.Hidden = False Then
                                script = script & vbCrLf & "<server>"
                                script = script & vbCrLf & "<properties>"
                                script = script & vbCrLf & "<displayName>" & cellServer.Value & " (" & cellDesc.Value & ")</displayName>"
                                script = script & vbCrLf & "<name>" & cellServer.Value & "</name>"
                                script = script & vbCrLf & "</properties>"
                                script = script & vbCrLf & "</server>"
                            End If
                            cnt = cnt + 1
                        Next
                    End If
                Next
                script = script & vbCrLf & "</file>"
                script = script & vbCrLf & "<connected />"
                script = script & vbCrLf & "<favorites />"
                script = script & vbCrLf & "<recentlyUsed />"
                script = script & vbCrLf & "</RDCMan>"

                'Debug.Print script
                'Close 'Close any open text files
                'nDestFile = FreeFile() 'Get the number of the next free text file
                'Open FileName For Output As #nDestFile 'Write the entire file to sText
                'Print #nDestFile, script
                System.IO.File.WriteAllText(FileName, script)

            Catch ex As Exception
                Call DisplayMessage(ex)

            Finally
                lstCol = Nothing
                tbl = Nothing
                col = Nothing
                cellServer = Nothing
                cellDesc = Nothing
                'Close

            End Try

        End Sub

        Public Sub RefreshCombobox(ByVal control As Office.IRibbonControl)
            Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveCell.ListObject
            Try
                If (tbl Is Nothing) Then
                    MessageBox.Show("Please select a table.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Try
                End If
                ribbon.Invalidate()

            Catch ex As Exception
                Call DisplayMessage(ex)

            Finally
                tbl = Nothing

            End Try

        End Sub

        ''' <summary>
        ''' show the settings form
        ''' </summary>
        ''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
        ''' <remarks></remarks>
        Public Sub OpenSettingsForm(ByVal control As Office.IRibbonControl)
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
                Call DisplayMessage(ex)

            End Try

        End Sub

        ''' <summary>
        ''' show the read me file
        ''' </summary>
        ''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
        ''' <remarks></remarks>
        Public Sub OpenHelpAsBuiltFile(ByVal control As Office.IRibbonControl)
            Try
                Call OpenFile(My.Settings.App_PathReadMe)

            Catch ex As Exception
                Call DisplayMessage(ex)

            End Try

        End Sub

#End Region

#Region "| Subroutines |"

        Public Function GetItem(col As Object, key As Object) As Object
            Try
                GetItem = col(key)

            Catch ex As Exception
                Call DisplayMessage(ex)
                GetItem = Nothing

            End Try

        End Function

        Public Function GetPingResult(Host) As String
            Dim objPing As Object
            Dim objStatus As Object
            Dim result As String = String.Empty
            Try

                objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").
                    ExecQuery("Select * from Win32_PingStatus Where Address = '" & Host & "'")

                For Each objStatus In objPing
                    Select Case objStatus.StatusCode
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
                Return result & " : " & Format(Now(), "dd-MMM-yyyy hh:nn:ss")

            Catch ex As Exception
                Call DisplayMessage(ex)
                Return "Error: " & ex.ToString()

            Finally
                objPing = Nothing
                objStatus = Nothing

            End Try

        End Function

        ''' <summary>
        ''' open a file from the source list
        ''' </summary>
        ''' <param name="fileName">The selected file to open</param>
        ''' <remarks></remarks>
        Public Sub OpenFile(ByVal fileName As String)
            Dim pStart As New System.Diagnostics.Process
            Try
                If fileName = String.Empty Then Exit Try
                pStart.StartInfo.FileName = fileName
                pStart.Start()

            Catch ex As System.ComponentModel.Win32Exception
                'MessageBox.Show("No application is assicated to this file type." & vbCrLf & vbCrLf & pstrFile, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try

            Catch ex As Exception
                Call DisplayMessage(ex)
                Exit Try

            Finally
                pStart.Dispose()

            End Try

        End Sub

#End Region

    End Class

End Namespace