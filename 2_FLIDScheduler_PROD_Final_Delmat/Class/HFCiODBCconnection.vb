Imports System.IO
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Configuration
Imports System.Collections.Specialized

Public Class HFCiODBCconnection
    Private moDbConn As New SqlConnection

    Public Sub OpenConnection(ByVal DBServer As String, ByVal DBUserName As String, ByVal DBPassword As String, ByVal DBName As String)
        Dim lsConnectionString As String

        Try
            'lsConnectionString = "Server=" & DBServer & ";Database=" & DBName & ";Uid=" & DBUserName & ";Pwd='" & DBPassword & "';MultipleActiveResultSets=True"
            'lsConnectionString = "Server=" & DBServer & ";Database=" & DBName & ";Integrated Security=True;MultipleActiveResultSets=True"
            'lsConnectionString = "Server=" & DBServer & ";Database=" & DBName & ";Integrated Security=SSPI;MultipleActiveResultSets=True"

            lsConnectionString = ConfigurationManager.AppSettings("HFCConnectionString")

            moDbConn = New SqlConnection(lsConnectionString)
            moDbConn.Open()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
    End Sub

    Public Sub CloseConnection()
        moDbConn.Close()
    End Sub

    Public Function InitiateTransaction() As SqlTransaction
        InitiateTransaction = moDbConn.BeginTransaction
    End Function

    Public Function GetDataReader(ByVal SQL As String) As SqlDataReader
        GetDataReader = Nothing
        Dim lobjCommand As New SqlCommand
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandText = SQL
            lobjCommand.CommandType = CommandType.Text
            GetDataReader = lobjCommand.ExecuteReader
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
        lobjCommand.Dispose()
        lobjCommand = Nothing
    End Function

    Public Function GetCurrentDate() As Date
        GetCurrentDate = Format(CDate(GetExecuteScalar("Select Now()")), "dd-MMM-yyyy hh:mm:ss")
    End Function
    Public Function GetDataReader(ByVal SQL As String, ByRef ObjTransaction As SqlTransaction) As SqlDataReader
        GetDataReader = Nothing
        Dim lobjCommand As New SqlCommand
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandText = SQL
            lobjCommand.CommandType = CommandType.Text
            lobjCommand.Transaction = ObjTransaction
            GetDataReader = lobjCommand.ExecuteReader
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
        lobjCommand.Dispose()
        lobjCommand = Nothing
    End Function

    Public Function GetExecuteScalar(ByVal SQL As String) As String
        GetExecuteScalar = ""
        Dim lobjCommand As New SqlCommand
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandText = SQL
            lobjCommand.CommandType = CommandType.Text
            GetExecuteScalar = lobjCommand.ExecuteScalar() & ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
        lobjCommand.Dispose()
        lobjCommand = Nothing
    End Function

    Public Function GetExecuteScalar(ByVal SQL As String, ByRef ObjTransaction As SqlTransaction) As String
        GetExecuteScalar = ""
        Dim lobjCommand As New SqlCommand
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandText = SQL
            lobjCommand.CommandType = CommandType.Text
            lobjCommand.Transaction = ObjTransaction
            GetExecuteScalar = lobjCommand.ExecuteScalar() & ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
        lobjCommand.Dispose()
        lobjCommand = Nothing
    End Function

    Public Function GetDataSet(ByVal SQL As String, ByVal TableName As String) As DataSet
        GetDataSet = Nothing
        Dim lobjDataAdapter As New SqlDataAdapter
        Dim lobjCommand As New SqlCommand
        Dim lobjDataSet As New DataSet
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandType = CommandType.Text
            lobjCommand.CommandText = SQL
            lobjDataAdapter.SelectCommand = lobjCommand
            lobjDataAdapter.Fill(lobjDataSet, TableName)
            Return lobjDataSet
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
        lobjDataSet.Dispose()
        lobjCommand.Dispose()
        lobjDataAdapter.Dispose()
        lobjDataSet = Nothing
        lobjCommand = Nothing
        lobjDataAdapter = Nothing
    End Function

    Public Function GetDataSet(ByVal SQL As String, ByVal TableName As String, ByRef ObjTransaction As SqlTransaction) As DataSet
        GetDataSet = Nothing
        Dim lobjDataAdapter As New SqlDataAdapter
        Dim lobjCommand As New SqlCommand
        Dim lobjDataSet As New DataSet
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandType = CommandType.Text
            lobjCommand.CommandText = SQL
            lobjDataAdapter.SelectCommand = lobjCommand
            lobjCommand.Transaction = ObjTransaction
            lobjDataAdapter.Fill(lobjDataSet, TableName)
            GetDataSet = lobjDataSet
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
        lobjDataSet.Dispose()
        lobjCommand.Dispose()
        lobjDataAdapter.Dispose()
        lobjDataSet = Nothing
        lobjCommand = Nothing
        lobjDataAdapter = Nothing
    End Function

    Public Function GetDataTable(ByVal SQL As String) As DataTable
        GetDataTable = Nothing
        Dim lobjDataAdapter As New SqlDataAdapter
        Dim lobjCommand As New SqlCommand
        Dim lobjDataTable As New Data.DataTable
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandType = CommandType.Text
            lobjCommand.CommandText = SQL
            lobjDataAdapter.SelectCommand = lobjCommand
            lobjDataAdapter.Fill(lobjDataTable)
            Return lobjDataTable
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
        lobjDataTable.Dispose()
        lobjCommand.Dispose()
        lobjDataAdapter.Dispose()
        lobjDataTable = Nothing
        lobjCommand = Nothing
        lobjDataAdapter = Nothing
    End Function

    Public Function GetDataTable(ByVal SQL As String, ByRef ObjTransaction As SqlTransaction) As DataTable
        GetDataTable = Nothing
        Dim lobjDataAdapter As New SqlDataAdapter
        Dim lobjCommand As New SqlCommand
        Dim lobjDataTable As New Data.DataTable
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandType = CommandType.Text
            lobjCommand.CommandText = SQL
            lobjDataAdapter.SelectCommand = lobjCommand
            lobjCommand.Transaction = ObjTransaction
            lobjDataAdapter.Fill(lobjDataTable)
            GetDataTable = lobjDataTable
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
        lobjDataTable.Dispose()
        lobjCommand.Dispose()
        lobjDataAdapter.Dispose()
        lobjDataTable = Nothing
        lobjCommand = Nothing
        lobjDataAdapter = Nothing
    End Function

    Public Function IsExists(ByVal SQL As String) As Boolean
        Dim lobjCommand As New SqlCommand
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandText = SQL
            lobjCommand.CommandType = CommandType.Text
            IsExists = IIf(Trim(lobjCommand.ExecuteScalar & "") = "", False, True)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
            IsExists = False
        End Try
        lobjCommand.Dispose()
        lobjCommand = Nothing
    End Function

    Public Function IsExists(ByVal SQL As String, ByRef ObjTransaction As SqlTransaction) As Boolean
        Dim lobjCommand As New SqlCommand
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandText = SQL
            lobjCommand.CommandType = CommandType.Text
            lobjCommand.Transaction = ObjTransaction
            IsExists = IIf(Trim(lobjCommand.ExecuteScalar & "") = "", False, True)
        Catch ex As Exception
            IsExists = False
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
        lobjCommand.Dispose()
        lobjCommand = Nothing
    End Function


    Public Function ExecuteNonQuerySQLCount(ByVal psSQL As String) As Integer
        Dim lobjCommand As New SqlCommand
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandText = psSQL
            lobjCommand.CommandType = CommandType.Text
            ExecuteNonQuerySQLCount = lobjCommand.ExecuteNonQuery()
        Catch ex As Exception
            ExecuteNonQuerySQLCount = -1
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
        lobjCommand.Dispose()
        lobjCommand = Nothing
    End Function

    Public Function ExecuteNonQuerySQLCount(ByVal psSQL As String, ByRef pObjTran As SqlTransaction) As Integer
        Dim lobjCommand As New SqlCommand
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandText = psSQL
            lobjCommand.CommandType = CommandType.Text
            lobjCommand.Transaction = pObjTran
            ExecuteNonQuerySQLCount = lobjCommand.ExecuteNonQuery()
        Catch ex As Exception
            ExecuteNonQuerySQLCount = -1
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
        lobjCommand.Dispose()
        lobjCommand = Nothing
    End Function


    Public Function ExecuteNonQuerySQL(ByVal SQL As String) As Object
        Dim lobjCommand As New SqlCommand
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandText = SQL
            lobjCommand.CommandType = CommandType.Text
            lobjCommand.ExecuteNonQuery()
            ExecuteNonQuerySQL = ""
        Catch ex As SqlException
            ExecuteNonQuerySQL = ex.Message
        End Try
        lobjCommand.Dispose()
        lobjCommand = Nothing
    End Function

    Public Function ExecuteNonQuerySQL(ByVal SQL As String, ByRef ObjTransaction As SqlTransaction) As Object
        Dim lobjCommand As New SqlCommand
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandText = SQL
            lobjCommand.CommandType = CommandType.Text
            lobjCommand.Transaction = ObjTransaction
            lobjCommand.ExecuteNonQuery()
            ExecuteNonQuerySQL = ""
        Catch ex As SqlException
            ExecuteNonQuerySQL = ex.Message
        End Try
        lobjCommand.Dispose()
        lobjCommand = Nothing
    End Function

    Public Sub FillComboBox(ByVal SQL As String, ByVal sDisplayMember As String, _
                            ByVal sValueMember As String, ByRef cComboBox As ComboBox)
        Dim lobjDataAdapter As New SqlDataAdapter
        Dim lobjCommand As New SqlCommand
        Dim lobjDataTable As New Data.DataTable
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandType = CommandType.Text
            lobjCommand.CommandText = SQL
            lobjDataAdapter.SelectCommand = lobjCommand
            lobjDataAdapter.Fill(lobjDataTable)
            cComboBox.DataSource = lobjDataTable
            cComboBox.DisplayMember = sDisplayMember
            cComboBox.ValueMember = sValueMember
            cComboBox.SelectedIndex = -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
        lobjCommand.Dispose()
        lobjDataAdapter.Dispose()
        lobjCommand = Nothing
        lobjDataAdapter = Nothing
    End Sub

    Public Sub FillListBox(ByVal SQL As String, ByRef cListBox As ListBox)
        Dim lobjCommand As New SqlCommand
        Dim lobjDataReader As SqlDataReader
        lobjDataReader = Nothing
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandType = CommandType.Text
            lobjCommand.CommandText = SQL
            lobjDataReader = lobjCommand.ExecuteReader

            While lobjDataReader.Read
                cListBox.Items.Add(lobjDataReader(0))
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try

        lobjCommand.Dispose()
        lobjDataReader.Dispose()
        lobjCommand = Nothing
        lobjDataReader = Nothing
    End Sub

    Public Sub FillCheckedListBox(ByVal SQL As String, ByRef cCheckedListBox As CheckedListBox)
        Dim lobjCommand As New SqlCommand
        Dim lobjDataReader As SqlDataReader
        lobjDataReader = Nothing
        Try
            lobjCommand.Connection = moDbConn
            lobjCommand.CommandType = CommandType.Text
            lobjCommand.CommandText = SQL
            lobjDataReader = lobjCommand.ExecuteReader

            While lobjDataReader.Read
                cCheckedListBox.Items.Add(lobjDataReader(0))
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
        lobjCommand.Dispose()
        lobjDataReader.Dispose()
        lobjCommand = Nothing
        lobjDataReader = Nothing
    End Sub

    Public Sub FillDataGridView(ByRef ObjDataGridView As DataGridView, ByVal SQL As String, _
    ByVal TableName As String, ByVal ColumnWidth As String, ByVal ColumnAlignment As String)

        Dim lObjDataSet As New DataSet
        Dim lsColWidth As String()
        Dim lsColAlignment As String()
        Dim liCount As Integer

        Try
            ObjDataGridView.DataSource = Nothing
            lObjDataSet = GetDataSet(SQL, TableName)
            ObjDataGridView.DataSource = lObjDataSet.Tables(TableName)

            lsColWidth = Split(ColumnWidth, "|")
            lsColAlignment = Split(ColumnAlignment, "|")

            If UBound(lsColWidth) = UBound(lsColAlignment) Then
                For liCount = 0 To UBound(lsColWidth)
                    ObjDataGridView.Columns(liCount).HeaderText = lObjDataSet.Tables(TableName).Columns(liCount).ColumnName
                    ObjDataGridView.Columns(liCount).Width = lsColWidth(liCount)
                    ObjDataGridView.Columns(liCount).ReadOnly = True
                    If Trim(lsColAlignment(liCount).ToString) = 1 Then
                        ObjDataGridView.Columns(liCount).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                    ElseIf Trim(lsColAlignment(liCount).ToString) = 2 Then
                        ObjDataGridView.Columns(liCount).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    ElseIf Trim(lsColAlignment(liCount).ToString) = 3 Then
                        ObjDataGridView.Columns(liCount).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End If
                    ObjDataGridView.ColumnHeadersDefaultCellStyle.NullValue = ""
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
    End Sub

    Public Sub ExportDataGridToXL(ByVal ObjDataTable As DataTable, ByVal FileName As String)
        Dim liTotCol As Integer
        Dim liCol As Integer
        Dim llRow As Long
        Dim lsSheetName As String
        Dim liSubSheet As Integer
        Dim lsOutputFile As String
        Dim i As Long

        liSubSheet = 0
        Try
            liTotCol = ObjDataTable.Columns.Count
            lsSheetName = "Report"
            lsOutputFile = FileName & ".xls"
            If File.Exists(lsOutputFile) Then File.Delete(lsOutputFile)
            FileOpen(1, lsOutputFile, OpenMode.Output)
            PrintLine(1, "<?xml version=""1.0"" encoding=""utf-8""?>")
            PrintLine(1, "<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:html=""http://www.w3.org/TR/REC-html40"">")
            PrintLine(1, "<Styles>")
            PrintLine(1, "<Style ss:ID=""s1"">")
            PrintLine(1, "<Font x:Family=""Swiss"" ss:Bold=""1""/>")
            PrintLine(1, "<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/>")
            PrintLine(1, "<Borders>")
            PrintLine(1, "<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
            PrintLine(1, "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
            PrintLine(1, "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
            PrintLine(1, "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
            PrintLine(1, "</Borders>")
            PrintLine(1, "</Style>")
            PrintLine(1, "<Style ss:ID=""s2"">")
            PrintLine(1, "<Borders>")
            PrintLine(1, "<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
            PrintLine(1, "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
            PrintLine(1, "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
            PrintLine(1, "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
            PrintLine(1, "</Borders>")
            PrintLine(1, "</Style>")
            PrintLine(1, "</Styles>")

            PrintLine(1, "<Worksheet ss:Name=""" & lsSheetName & """>")
            PrintLine(1, "<Table>")

            PrintLine(1, "<Row>")
            For liCol = 1 To liTotCol
                PrintLine(1, "<Cell ss:StyleID=""s1""><Data ss:Type=""String"">" & ObjDataTable.Columns(liCol - 1).Caption.ToString & "</Data></Cell>")
            Next
            PrintLine(1, "</Row>")
            For llRow = 0 To ObjDataTable.Rows.Count - 1
                If i > 65000 Then
                    PrintLine(1, "</Table>")
                    PrintLine(1, "</Worksheet>")
                    PrintLine(1, "</Workbook>")
                    FileClose(1)
                    i = 1
                    liSubSheet = liSubSheet + 1
                    FileName = FileName & liSubSheet
                    lsSheetName = "Report"
                    lsOutputFile = FileName & ".xls"
                    If File.Exists(lsOutputFile) Then File.Delete(lsOutputFile)
                    FileOpen(1, lsOutputFile, OpenMode.Output)
                    PrintLine(1, "<?xml version=""1.0"" encoding=""utf-8""?>")
                    PrintLine(1, "<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:html=""http://www.w3.org/TR/REC-html40"">")
                    PrintLine(1, "<Styles>")
                    PrintLine(1, "<Style ss:ID=""s1"">")
                    PrintLine(1, "<Font x:Family=""Swiss"" ss:Bold=""1""/>")
                    PrintLine(1, "<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/>")
                    PrintLine(1, "<Borders>")
                    PrintLine(1, "<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
                    PrintLine(1, "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
                    PrintLine(1, "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
                    PrintLine(1, "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
                    PrintLine(1, "</Borders>")
                    PrintLine(1, "</Style>")
                    PrintLine(1, "<Style ss:ID=""s2"">")
                    PrintLine(1, "<Borders>")
                    PrintLine(1, "<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
                    PrintLine(1, "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
                    PrintLine(1, "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
                    PrintLine(1, "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
                    PrintLine(1, "</Borders>")
                    PrintLine(1, "</Style>")
                    PrintLine(1, "</Styles>")

                    PrintLine(1, "<Worksheet ss:Name=""" & lsSheetName & """>")
                    PrintLine(1, "<Table>")

                    PrintLine(1, "<Row>")
                    For liCol = 1 To liTotCol
                        PrintLine(1, "<Cell ss:StyleID=""s1""><Data ss:Type=""String"">" & ObjDataTable.Columns(liCol - 1).Caption.ToString & "</Data></Cell>")
                    Next
                    PrintLine(1, "</Row>")
                End If
                PrintLine(1, "<Row>")
                For liCol = 1 To liTotCol
                    PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & ObjDataTable.Rows(llRow).Item(liCol - 1).ToString & "</Data></Cell>")
                Next
                PrintLine(1, "</Row>")
                i = i + 1
            Next
            PrintLine(1, "</Table>")
            PrintLine(1, "</Worksheet>")
            PrintLine(1, "</Workbook>")
            FileClose(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
    End Sub

    Public Sub ExportSqlToXL(ByRef FileName As String, ByRef SheetCount As Short, ByRef SheetName As Object, ByRef SQL As Object)
        Dim lObjDataTable As New DataTable
        Dim liTotCol As Integer
        Dim liCol As Integer
        Dim llRow As Long

        Dim lsName As String
        Dim liSubSheet As Integer
        Dim lsOutputFile As String

        Dim liShtNo As Integer
        Dim lsQry As String

        Dim llRecCnt As Long

        liSubSheet = 0

        Try
            lsOutputFile = FileName & ".xls"
            If File.Exists(lsOutputFile) Then File.Delete(lsOutputFile)
            FileOpen(1, lsOutputFile, OpenMode.Output)
            PrintLine(1, "<?xml version=""1.0"" encoding=""utf-8""?>")
            PrintLine(1, "<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:html=""http://www.w3.org/TR/REC-html40"">")
            PrintLine(1, "<Styles>")
            PrintLine(1, "<Style ss:ID=""s1"">")
            PrintLine(1, "<Font x:Family=""Swiss"" ss:Bold=""1""/>")
            PrintLine(1, "<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/>")
            PrintLine(1, "<Borders>")
            PrintLine(1, "<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
            PrintLine(1, "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
            PrintLine(1, "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
            PrintLine(1, "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
            PrintLine(1, "</Borders>")
            PrintLine(1, "</Style>")
            PrintLine(1, "<Style ss:ID=""s2"">")
            PrintLine(1, "<Borders>")
            PrintLine(1, "<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
            PrintLine(1, "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
            PrintLine(1, "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
            PrintLine(1, "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
            PrintLine(1, "</Borders>")
            PrintLine(1, "</Style>")
            PrintLine(1, "</Styles>")
            For liShtNo = 0 To SheetCount - 1
                lsQry = SQL(liShtNo)
                lObjDataTable = GetDataTable(lsQry)
                liTotCol = lObjDataTable.Columns.Count
                PrintLine(1, "<Worksheet ss:Name=""" & SheetName(liShtNo) & """>")
                PrintLine(1, "<Table>")

                PrintLine(1, "<Row>")
                For liCol = 1 To liTotCol
                    PrintLine(1, "<Cell ss:StyleID=""s1""><Data ss:Type=""String"">" & lObjDataTable.Columns(liCol - 1).Caption.ToString & "</Data></Cell>")
                Next
                PrintLine(1, "</Row>")

                For llRow = 0 To lObjDataTable.Rows.Count - 1
                    If llRecCnt > 65000 Then
                        llRecCnt = 1
                        liSubSheet = liSubSheet + 1
                        lsName = SheetName(liShtNo) & liSubSheet
                        PrintLine(1, "</Table>")
                        PrintLine(1, "</Worksheet>")
                        PrintLine(1, "<Worksheet ss:Name=""" & lsName & """>")
                        PrintLine(1, "<Table>")

                        PrintLine(1, "<Row>")
                        For liCol = 1 To liTotCol
                            PrintLine(1, "<Cell ss:StyleID=""s1""><Data ss:Type=""String"">" & lObjDataTable.Columns(liCol - 1).Caption.ToString & "</Data></Cell>")
                        Next
                        PrintLine(1, "</Row>")
                    End If
                    PrintLine(1, "<Row>")
                    For liCol = 1 To liTotCol
                        PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & lObjDataTable.Rows(llRow).Item(liCol - 1).ToString & "</Data></Cell>")
                    Next
                    PrintLine(1, "</Row>")
                    llRecCnt = llRecCnt + 1
                Next
                PrintLine(1, "</Table>")
                PrintLine(1, "</Worksheet>")
            Next
            PrintLine(1, "</Workbook>")
            FileClose(1)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "i-Alert")
        End Try
    End Sub

    Public Sub AutoTextCombo(ByVal ObjSqlDataReader As SqlDataReader, ByVal cComboBox As ComboBox, _
                             ByVal KeyInFlag As Boolean, ByVal KeyInTxt As String, ByVal PrevTxt As String)
        With cComboBox
            If Trim(.Text) = "" Then Exit Sub
            If KeyInFlag = True Then
                If UCase(.Text) <> UCase(PrevTxt) Then
                    KeyInTxt = .Text
                    PrevTxt = KeyInTxt
                    If ObjSqlDataReader.HasRows = True Then
                        While ObjSqlDataReader.HasRows = True
                            ObjSqlDataReader.Read()
                            If UCase(.Text) = UCase(Mid(ObjSqlDataReader(0), 1, Len(.Text))) Then
                                KeyInFlag = False
                                .Text = ObjSqlDataReader(0)
                                .SelectionStart = Len(KeyInTxt)
                                .SelectionLength = Len(.Text) - Len(KeyInTxt)
                                Exit Sub
                            End If
                        End While
                    End If
                End If
            Else
                KeyInFlag = True
            End If
        End With
    End Sub
End Class
