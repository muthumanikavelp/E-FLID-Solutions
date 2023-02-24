Imports System.Data.Sql
Imports System.IO
Imports System.Text
Imports System.Drawing

Module modExcel

    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, _
            ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
            ByVal nShowCmd As Integer) As Integer
    Const SW_SHOWNORMAL As Short = 1

    ' ''Public Sub DgridToExcel(ByVal Dgrid As DataGrid, ByVal ExcelFileName As String, ByVal SheetName As String, Optional ByVal StoreasNum As String = "")

    ' ''    Dim Col As Integer, Row As Integer, TotCol As Integer, TotSheets As Integer
    ' ''    Dim NumericCols() As String
    ' ''    Dim objApplication As New Excel.Application
    ' ''    Dim objBooks As Excel.Workbooks
    ' ''    Dim objWorkBook As Excel.Workbook
    ' ''    Dim objWorkSheet As Excel.Worksheet

    ' ''    Try
    ' ''        NumericCols = Split(StoreasNum, "|")

    ' ''        Row = 1

    ' ''        If File.Exists(ExcelFileName) = False Then
    ' ''            objBooks = objApplication.Workbooks
    ' ''            objWorkBook = objBooks.Add()
    ' ''            objWorkSheet = objApplication.Worksheets.Add()
    ' ''            objWorkSheet.Name = SheetName
    ' ''            objWorkBook.SaveAs(ExcelFileName, 1)
    ' ''        Else
    ' ''            objBooks = objApplication.Workbooks
    ' ''            objWorkBook = objBooks.Open(ExcelFileName, False, False)
    ' ''            TotSheets = objWorkBook.Sheets.Count
    ' ''            While TotSheets > 0
    ' ''                objWorkSheet = objWorkBook.Sheets(TotSheets)
    ' ''                If objWorkSheet.Name = SheetName Then
    ' ''                    SheetName = SheetName & "+"
    ' ''                    TotSheets = objWorkBook.Sheets.Count
    ' ''                Else
    ' ''                    TotSheets = TotSheets - 1
    ' ''                End If
    ' ''            End While
    ' ''            objWorkSheet = objApplication.Worksheets.Add()
    ' ''            objWorkSheet.Name = SheetName
    ' ''        End If

    ' ''        objApplication.Visible = True

    ' ''        'objDataTable = Dgrid.DataSource
    ' ''        TotCol = Dgrid.VisibleColumnCount

    ' ''        For Col = 1 To TotCol
    ' ''            objWorkSheet.Cells(Row, Col) = Dgrid.Item(0, Col - 1).ToString
    ' ''            objWorkSheet.Cells(Row, Col).interior.colorindex = 10
    ' ''            objWorkSheet.Cells(Row, Col).font.colorindex = 6
    ' ''            objWorkSheet.Cells(Row, Col).Font.Bold = True
    ' ''            objWorkSheet.Cells(Row, Col).Borders.ColorIndex = 56
    ' ''        Next


    ' ''        For Col = 1 To TotCol
    ' ''            For Row = 1 To Dgrid.VisibleRowCount
    ' ''                If IsNumericFldCol(Col, NumericCols) = False Then
    ' ''                    objWorkSheet.Cells(Row + 2, Col) = "'" & Dgrid.Item(Col - 1, Row)
    ' ''                Else
    ' ''                    objWorkSheet.Cells(Row + 2, Col) = Dgrid.Item(Col - 1, Row).Value.ToString
    ' ''                End If
    ' ''                objWorkSheet.Cells(Row + 2, Col).Borders.ColorIndex = 56
    ' ''            Next
    ' ''        Next

    ' ''        objWorkSheet.Columns.AutoFit()
    ' ''        For Col = 1 To TotCol
    ' ''            If objWorkSheet.Columns(Col).ColumnWidth > 40 Then
    ' ''                objWorkSheet.Columns(Col).ColumnWidth = 40
    ' ''                objWorkSheet.Columns(Col).WrapText = True
    ' ''            End If
    ' ''        Next


    ' ''        objWorkBook.Save()

    ' ''        'NAR(objWorkSheet)
    ' ''        objWorkBook.Close(False)
    ' ''        'NAR(objWorkBook)
    ' ''        objBooks.Close()
    ' ''        'NAR(objBooks)
    ' ''        objApplication.Quit()
    ' ''        'NAR(objApplication)
    ' ''        'GC.Collect()
    ' ''        'GC.WaitForPendingFinalizers()
    ' ''        MsgBox("Created Successfully ", MsgBoxStyle.Information)
    ' ''    Catch ex As Exception
    ' ''        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error...")
    ' ''    End Try
    ' ''End Sub
    ' ''Public Sub DgridViewToExcel(ByVal Dgrid As DataGridView, ByVal ExcelFileName As String, ByVal SheetName As String, Optional ByVal StoreasNum As String = "")

    ' ''    Dim Col As Integer, Row As Integer, TotCol As Integer, TotSheets As Integer
    ' ''    Dim NumericCols() As String
    ' ''    'Dim objDataTable As New DataGridView
    ' ''    Dim objApplication As New Excel.Application
    ' ''    Dim objBooks As Excel.Workbooks
    ' ''    Dim objWorkBook As Excel.Workbook
    ' ''    Dim objWorkSheet As Excel.Worksheet

    ' ''    Try
    ' ''        NumericCols = Split(StoreasNum, "|")

    ' ''        Row = 1

    ' ''        If File.Exists(ExcelFileName) = False Then
    ' ''            objBooks = objApplication.Workbooks
    ' ''            objWorkBook = objBooks.Add()
    ' ''            objWorkSheet = objApplication.Worksheets.Add()
    ' ''            objWorkSheet.Name = SheetName
    ' ''            objWorkBook.SaveAs(ExcelFileName, 1)
    ' ''        Else
    ' ''            objBooks = objApplication.Workbooks
    ' ''            objWorkBook = objBooks.Open(ExcelFileName, False, False)
    ' ''            TotSheets = objWorkBook.Sheets.Count
    ' ''            While TotSheets > 0
    ' ''                objWorkSheet = objWorkBook.Sheets(TotSheets)
    ' ''                If objWorkSheet.Name = SheetName Then
    ' ''                    SheetName = SheetName & "+"
    ' ''                    TotSheets = objWorkBook.Sheets.Count
    ' ''                Else
    ' ''                    TotSheets = TotSheets - 1
    ' ''                End If
    ' ''            End While
    ' ''            objWorkSheet = objApplication.Worksheets.Add()
    ' ''            objWorkSheet.Name = SheetName
    ' ''        End If

    ' ''        objApplication.Visible = True

    ' ''        'objDataTable = Dgrid.DataSource
    ' ''        TotCol = Dgrid.ColumnCount

    ' ''        For Col = 1 To TotCol
    ' ''            objWorkSheet.Cells(Row, Col) = Dgrid.Columns(Col - 1).Name.ToString
    ' ''            objWorkSheet.Cells(Row, Col).interior.colorindex = 10
    ' ''            objWorkSheet.Cells(Row, Col).font.colorindex = 6
    ' ''            objWorkSheet.Cells(Row, Col).Font.Bold = True
    ' ''            objWorkSheet.Cells(Row, Col).Borders.ColorIndex = 56
    ' ''        Next

    ' ''        For Row = 0 To Dgrid.RowCount - 1
    ' ''            For Col = 1 To TotCol
    ' ''                If IsNumericFldCol(Col, NumericCols) = False Then
    ' ''                    objWorkSheet.Cells(Row + 2, Col) = "'" & Dgrid.Item(Col - 1, Row).Value.ToString
    ' ''                Else
    ' ''                    objWorkSheet.Cells(Row + 2, Col) = Dgrid.Item(Col - 1, Row).Value.ToString
    ' ''                End If
    ' ''                objWorkSheet.Cells(Row + 2, Col).Borders.ColorIndex = 56
    ' ''            Next
    ' ''        Next

    ' ''        objWorkSheet.Columns.AutoFit()
    ' ''        For Col = 1 To TotCol
    ' ''            If objWorkSheet.Columns(Col).ColumnWidth > 40 Then
    ' ''                objWorkSheet.Columns(Col).ColumnWidth = 40
    ' ''                objWorkSheet.Columns(Col).WrapText = True
    ' ''            End If
    ' ''        Next


    ' ''        objWorkBook.Save()

    ' ''        'NAR(objWorkSheet)
    ' ''        objWorkBook.Close(False)
    ' ''        'NAR(objWorkBook)
    ' ''        objBooks.Close()
    ' ''        'NAR(objBooks)
    ' ''        objApplication.Quit()
    ' ''        'NAR(objApplication)
    ' ''        'GC.Collect()
    ' ''        'GC.WaitForPendingFinalizers()
    ' ''        MsgBox("Created Successfully ", MsgBoxStyle.Information)
    ' ''    Catch ex As Exception
    ' ''        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error...")
    ' ''    End Try
    ' ''End Sub

    ' ''Public Sub SqlToExcel(ByVal SqlStr As String, ByVal ExcelFileName As String, ByVal SheetName As String, Optional ByVal StoreasNum As String = "")

    ' ''    Dim objFile As File
    ' ''    Dim Col As Integer, Row As Integer, TotSheets As Integer
    ' ''    Dim NumericCols() As String

    ' ''    Dim objDataReader As Odbc.OdbcDataReader
    ' ''    Dim objApplication As New Excel.Application
    ' ''    Dim objBooks As Excel.Workbooks
    ' ''    Dim objWorkBook As Excel.Workbook
    ' ''    Dim objWorkSheet As Excel.Worksheet

    ' ''    Try

    ' ''        objDataReader = gfExecuteQry(SqlStr, gOdbcConn)
    ' ''        If objDataReader.HasRows = False Then
    ' ''            MsgBox("No records found...", MsgBoxStyle.Information, gProjectName)
    ' ''            Exit Sub
    ' ''        End If

    ' ''        Row = 0
    ' ''        NumericCols = Split(StoreasNum, "|")

    ' ''        If objFile.Exists(ExcelFileName) = False Then
    ' ''            objBooks = objApplication.Workbooks
    ' ''            objWorkBook = objBooks.Add()
    ' ''            objWorkSheet = objApplication.Worksheets.Add()
    ' ''            objWorkSheet.Name = SheetName
    ' ''            objWorkBook.SaveAs(ExcelFileName, 1)
    ' ''        Else
    ' ''            objBooks = objApplication.Workbooks
    ' ''            objWorkBook = objBooks.Open(ExcelFileName, False, False)
    ' ''            TotSheets = objWorkBook.Sheets.Count
    ' ''            While TotSheets > 0
    ' ''                objWorkSheet = objWorkBook.Sheets(TotSheets)
    ' ''                If objWorkSheet.Name = SheetName Then
    ' ''                    SheetName = SheetName & "+"
    ' ''                    TotSheets = objWorkBook.Sheets.Count
    ' ''                Else
    ' ''                    TotSheets = TotSheets - 1
    ' ''                End If
    ' ''            End While
    ' ''            objWorkSheet = objApplication.Worksheets.Add()
    ' ''            objWorkSheet.Name = SheetName
    ' ''        End If

    ' ''        objApplication.Visible = True

    ' ''        'objDataReader = gfExecuteQry(SqlStr, gOdbcConn)
    ' ''        If objDataReader.HasRows = True Then
    ' ''            Row = Row + 1
    ' ''            For Col = 0 To objDataReader.FieldCount - 1
    ' ''                objWorkSheet.Cells(Row, Col + 1) = "'" & objDataReader.GetName(Col).ToString
    ' ''                objWorkSheet.Cells(Row, Col + 1).interior.colorindex = 10
    ' ''                objWorkSheet.Cells(Row, Col + 1).font.colorindex = 6
    ' ''                objWorkSheet.Cells(Row, Col + 1).font.bold = True
    ' ''                objWorkSheet.Cells(Row, Col + 1).Borders.ColorIndex = 56
    ' ''            Next
    ' ''            While objDataReader.Read()
    ' ''                Row = Row + 1
    ' ''                With objWorkSheet
    ' ''                    For Col = 0 To objDataReader.FieldCount - 1
    ' ''                        If IsNumericFldCol(Col + 1, NumericCols) = False Then
    ' ''                            .Cells(Row, Col + 1) = "'" & objDataReader.Item(Col).ToString
    ' ''                        Else
    ' ''                            .Cells(Row, Col + 1) = objDataReader.Item(Col).ToString
    ' ''                        End If
    ' ''                    Next
    ' ''                End With
    ' ''            End While
    ' ''        End If
    ' ''        objDataReader.Close()

    ' ''        objWorkSheet.Columns.AutoFit()
    ' ''        For Col = 0 To objDataReader.FieldCount - 1
    ' ''            If objWorkSheet.Columns(Col + 1).ColumnWidth > 40 Then
    ' ''                objWorkSheet.Columns(Col + 1).ColumnWidth = 40
    ' ''                objWorkSheet.Columns(Col + 1).WrapText = True
    ' ''            End If
    ' ''        Next

    ' ''        objWorkBook.Save()

    ' ''        'NAR(objWorkSheet)
    ' ''        objWorkBook.Close(False)
    ' ''        'NAR(objWorkBook)
    ' ''        objBooks.Close()
    ' ''        'NAR(objBooks)
    ' ''        objApplication.Quit()
    ' ''        'NAR(objApplication)
    ' ''        'GC.Collect()
    ' ''        'GC.WaitForPendingFinalizers()
    ' ''        MsgBox("Created Successfully ", MsgBoxStyle.Information)
    ' ''    Catch ex As Exception
    ' ''        'objMail.GF_Mail(gsMailFrom, gsMailTo, gsMailSub, ex.Message, gsFormName, "SqlToExcel")
    ' ''        MsgBox(ex.Message)
    ' ''    End Try
    ' ''End Sub

    ' ''Public Sub SqlToExcel1(ByVal SqlStr As String, ByVal ExcelFileName As String, ByVal SheetName As String, Optional ByVal StoreasNum As String = "")

    ' ''    Dim Col As Integer, Row As Integer, TotSheets As Integer
    ' ''    Dim NumericCols() As String

    ' ''    Dim objDataReader As Odbc.OdbcDataReader
    ' ''    Dim objApplication As New Excel.Application
    ' ''    Dim objBooks As Excel.Workbooks
    ' ''    Dim objWorkBook As Excel.Workbook
    ' ''    Dim objWorkSheet As Excel.Worksheet

    ' ''    Row = 0
    ' ''    NumericCols = Split(StoreasNum, "|")

    ' ''    If File.Exists(ExcelFileName) = False Then
    ' ''        objBooks = objApplication.Workbooks
    ' ''        objWorkBook = objBooks.Add()
    ' ''        objWorkSheet = objApplication.Worksheets.Add()
    ' ''        objWorkSheet.Name = SheetName
    ' ''        objWorkBook.SaveAs(ExcelFileName, 1)
    ' ''    Else
    ' ''        objBooks = objApplication.Workbooks
    ' ''        objWorkBook = objBooks.Open(ExcelFileName, False, False)
    ' ''        TotSheets = objWorkBook.Sheets.Count
    ' ''        While TotSheets > 0
    ' ''            objWorkSheet = objWorkBook.Sheets(TotSheets)
    ' ''            If objWorkSheet.Name = SheetName Then
    ' ''                SheetName = SheetName & "+"
    ' ''                TotSheets = objWorkBook.Sheets.Count
    ' ''            Else
    ' ''                TotSheets = TotSheets - 1
    ' ''            End If
    ' ''        End While
    ' ''        objWorkSheet = objApplication.Worksheets.Add()
    ' ''        objWorkSheet.Name = SheetName
    ' ''    End If

    ' ''    objApplication.Visible = True

    ' ''    objDataReader = gfExecuteQry(SqlStr, gOdbcConn)
    ' ''    If objDataReader.HasRows = True Then
    ' ''        Row = Row + 1
    ' ''        For Col = 0 To objDataReader.FieldCount - 1
    ' ''            objWorkSheet.Cells(Row, Col + 1) = "'" & objDataReader.GetName(Col).ToString
    ' ''            objWorkSheet.Cells(Row, Col + 1).interior.colorindex = 10
    ' ''            objWorkSheet.Cells(Row, Col + 1).font.colorindex = 6
    ' ''            objWorkSheet.Cells(Row, Col + 1).font.bold = True
    ' ''            objWorkSheet.Cells(Row, Col + 1).Borders.ColorIndex = 56
    ' ''        Next
    ' ''        While objDataReader.Read()
    ' ''            Row = Row + 1
    ' ''            With objWorkSheet
    ' ''                For Col = 0 To objDataReader.FieldCount - 1
    ' ''                    If IsNumericFldCol(Col + 1, NumericCols) = False Then
    ' ''                        .Cells(Row, Col + 1) = "'" & objDataReader.Item(Col).ToString
    ' ''                    Else
    ' ''                        .Cells(Row, Col + 1) = objDataReader.Item(Col).ToString
    ' ''                    End If
    ' ''                Next
    ' ''            End With
    ' ''        End While
    ' ''    End If
    ' ''    objDataReader.Close()

    ' ''    objWorkSheet.Columns.AutoFit()
    ' ''    For Col = 0 To objDataReader.FieldCount - 1
    ' ''        If objWorkSheet.Columns(Col + 1).ColumnWidth > 40 Then
    ' ''            objWorkSheet.Columns(Col + 1).ColumnWidth = 40
    ' ''            objWorkSheet.Columns(Col + 1).WrapText = True
    ' ''        End If
    ' ''    Next

    ' ''    objWorkBook.Save()

    ' ''    'NAR(objWorkSheet)
    ' ''    objWorkBook.Close(False)
    ' ''    'NAR(objWorkBook)
    ' ''    objBooks.Close()
    ' ''    'NAR(objBooks)
    ' ''    objApplication.Quit()
    ' ''    'NAR(objApplication)
    ' ''    'GC.Collect()
    ' ''    'GC.WaitForPendingFinalizers()
    ' ''    MsgBox("Created Successfully ", MsgBoxStyle.Information)
    ' ''End Sub
    Private Sub NAR(ByVal o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        Finally
            o = Nothing
        End Try
    End Sub
    ' ''Public Sub FlexToExcel(ByVal Fgrid As AxMSFlexGridLib.AxMSFlexGrid, ByVal ExcelFileName As String, ByVal SheetName As String, Optional ByVal StoreasNum As String = "")
    ' ''    Dim Col As Integer, Row As Integer, TotSheets As Integer
    ' ''    Dim NumericCols() As String
    ' ''    Dim objApplication As New Excel.Application
    ' ''    Dim objBooks As Excel.Workbooks
    ' ''    Dim objWorkBook As Excel.Workbook
    ' ''    Dim objWorkSheet As Excel.Worksheet

    ' ''    Row = 0
    ' ''    NumericCols = Split(StoreasNum, "|")

    ' ''    If File.Exists(ExcelFileName) = False Then
    ' ''        objBooks = objApplication.Workbooks
    ' ''        objWorkBook = objBooks.Add()
    ' ''        objWorkSheet = objApplication.Worksheets.Add()
    ' ''        objWorkSheet.Name = SheetName
    ' ''        objWorkBook.SaveAs(ExcelFileName, 1)
    ' ''    Else
    ' ''        objBooks = objApplication.Workbooks
    ' ''        objWorkBook = objBooks.Open(ExcelFileName, False, False)
    ' ''        TotSheets = objWorkBook.Sheets.Count
    ' ''        While TotSheets > 0
    ' ''            objWorkSheet = objWorkBook.Sheets(TotSheets)
    ' ''            If objWorkSheet.Name = SheetName Then
    ' ''                SheetName = SheetName & "+"
    ' ''                TotSheets = objWorkBook.Sheets.Count
    ' ''            Else
    ' ''                TotSheets = TotSheets - 1
    ' ''            End If
    ' ''        End While
    ' ''        objWorkSheet = objApplication.Worksheets.Add()
    ' ''        objWorkSheet.Name = SheetName
    ' ''    End If

    ' ''    objApplication.Visible = True
    ' ''    With Fgrid
    ' ''        For Row = 0 To .Rows - 1
    ' ''            .Row = Row
    ' ''            For Col = 0 To .Cols - 1
    ' ''                .Col = Col
    ' ''                If .CellBackColor.ToString <> Color.Black.ToString Then
    ' ''                    objWorkSheet.Cells(Row + 1, Col + 1).Interior.Color = ColorTranslator.ToWin32(.CellBackColor)
    ' ''                End If
    ' ''                If Row = 0 Then
    ' ''                    objWorkSheet.Cells(Row + 1, Col + 1).Font.Bold = True
    ' ''                End If
    ' ''                If IsNumericFldCol(Col + 1, NumericCols) = False Then
    ' ''                    objWorkSheet.Cells(Row + 1, Col + 1) = "'" & .get_TextMatrix(Row, Col)
    ' ''                Else
    ' ''                    objWorkSheet.Cells(Row + 1, Col + 1) = .get_TextMatrix(Row, Col)
    ' ''                End If
    ' ''                objWorkSheet.Cells(Row + 1, Col + 1).Borders.ColorIndex = 56
    ' ''            Next Col
    ' ''        Next Row
    ' ''        objWorkSheet.Columns.AutoFit()
    ' ''        For Col = 0 To .Cols - 1
    ' ''            If objWorkSheet.Columns(Col + 1).ColumnWidth > 40 Then
    ' ''                objWorkSheet.Columns(Col + 1).ColumnWidth = 40
    ' ''                objWorkSheet.Columns(Col + 1).WrapText = True
    ' ''            End If
    ' ''        Next
    ' ''    End With

    ' ''    objWorkBook.Save()

    ' ''    'NAR(objWorkSheet)
    ' ''    objWorkBook.Close(False)
    ' ''    'NAR(objWorkBook)
    ' ''    objBooks.Close()
    ' ''    'NAR(objBooks)
    ' ''    objApplication.Quit()
    ' ''    'NAR(objApplication)
    ' ''    'GC.Collect()
    ' ''    'GC.WaitForPendingFinalizers()
    ' ''    MsgBox("Created Successfully ", MsgBoxStyle.Information)
    ' ''End Sub
    ' ''Public Sub SeparatorToExcel(ByVal SourceFileName As String, ByVal ExcelFileName As String, ByVal SheetName As String, ByVal Separator As String, Optional ByVal StoreasNum As String = "")

    ' ''    Dim objStreamReader As StreamReader
    ' ''    Dim objApplication As New Excel.Application
    ' ''    Dim objBooks As Excel.Workbooks
    ' ''    Dim objWorkBook As Excel.Workbook
    ' ''    Dim objWorkSheet As Excel.Worksheet
    ' ''    Dim NumericCols() As String
    ' ''    Dim txt As String
    ' ''    Dim SplitTxt() As String
    ' ''    Dim lsRow As Long, lsCol As Integer, TotCol As Integer, TotSheets As Integer

    ' ''    NumericCols = Split(StoreasNum, "|")

    ' ''    If File.Exists(ExcelFileName) = False Then
    ' ''        objBooks = objApplication.Workbooks
    ' ''        objWorkBook = objBooks.Add()
    ' ''        objWorkSheet = objApplication.Worksheets.Add()
    ' ''        objWorkSheet.Name = SheetName
    ' ''        objWorkBook.SaveAs(ExcelFileName, 1)
    ' ''    Else
    ' ''        objBooks = objApplication.Workbooks
    ' ''        objWorkBook = objBooks.Open(ExcelFileName, False, False)
    ' ''        TotSheets = objWorkBook.Sheets.Count
    ' ''        While TotSheets > 0
    ' ''            objWorkSheet = objWorkBook.Sheets(TotSheets)
    ' ''            If objWorkSheet.Name = SheetName Then
    ' ''                SheetName = SheetName & "+"
    ' ''                TotSheets = objWorkBook.Sheets.Count
    ' ''            Else
    ' ''                TotSheets = TotSheets - 1
    ' ''            End If
    ' ''        End While
    ' ''        objWorkSheet = objApplication.Worksheets.Add()
    ' ''        objWorkSheet.Name = SheetName
    ' ''    End If

    ' ''    objApplication.Visible = True
    ' ''    objStreamReader = File.OpenText(SourceFileName.Trim)

    ' ''    While objStreamReader.Peek <> -1
    ' ''        txt = objStreamReader.ReadLine
    ' ''        SplitTxt = Split(txt, Separator)
    ' ''        TotCol = UBound(SplitTxt)
    ' ''        lsRow = lsRow + 1
    ' ''        For lsCol = 0 To TotCol
    ' ''            If IsNumericFldCol(lsCol + 1, NumericCols) = False Then
    ' ''                objWorkSheet.Cells(lsRow, lsCol + 1).value = "'" & CStr(SplitTxt(lsCol) & "")
    ' ''            Else
    ' ''                objWorkSheet.Cells(lsRow, lsCol + 1).value = CStr(SplitTxt(lsCol) & "")
    ' ''            End If
    ' ''            objWorkSheet.Cells(lsRow, lsCol + 1).Borders.ColorIndex = 56
    ' ''            objWorkSheet.Cells(lsRow, lsCol + 1).Show()
    ' ''        Next
    ' ''    End While

    ' ''    For lsCol = 1 To TotCol + 1
    ' ''        objWorkSheet.Columns(lsCol).Autofit()
    ' ''        If objWorkSheet.Columns(lsCol).ColumnWidth > 40 Then
    ' ''            objWorkSheet.Columns(lsCol).ColumnWidth = 40
    ' ''            objWorkSheet.Columns(lsCol).WrapText = True
    ' ''        End If
    ' ''    Next

    ' ''    objWorkBook.Save()
    ' ''    'NAR(objWorkSheet)
    ' ''    objWorkBook.Close(False)
    ' ''    'NAR(objWorkBook)
    ' ''    objBooks.Close()
    ' ''    'NAR(objBooks)
    ' ''    objApplication.Quit()
    ' ''    'NAR(objApplication)
    ' ''    'GC.Collect()
    ' ''    'GC.WaitForPendingFinalizers()
    ' ''    MsgBox("Created Successfully ", MsgBoxStyle.Information)
    ' ''End Sub
    Public Sub PrintDGridXML(ByVal dgrid As DataGridView, ByVal FileName As String, Optional ByVal SheetName As String = "Report", Optional ByVal StoreasNum As String = "", Optional ByVal DisplayFlag As Boolean = True)

        Dim objDataTable As New Data.DataTable
        Dim TotCol As Integer, Col As Integer, Row As Integer
        Dim objStreamReader As StreamReader
        Dim lsTextContent As String
        Dim NumericCols() As String
        NumericCols = Split(StoreasNum, "|")

        Try

            objDataTable = dgrid.DataSource
            TotCol = objDataTable.Columns.Count
            If File.Exists(FileName) Then File.Delete(FileName)
            If File.Exists(FileName) = False Then

                FileOpen(1, FileName, OpenMode.Output)

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

            Else
                objStreamReader = File.OpenText(FileName)
                lsTextContent = objStreamReader.ReadToEnd()
                lsTextContent = Replace(lsTextContent.ToString, "</Workbook>", "")
                objStreamReader.Close()
                FileOpen(1, FileName, OpenMode.Output)
                PrintLine(1, lsTextContent.ToString)
            End If

            PrintLine(1, "<Worksheet ss:Name=""" & SheetName & """>")
            PrintLine(1, "<Table>")

            PrintLine(1, "<Row>")
            For Col = 1 To TotCol
                PrintLine(1, "<Cell ss:StyleID=""s1""><Data ss:Type=""String"">" & objDataTable.Columns(Col - 1).Caption.ToString & "</Data></Cell>")
            Next
            PrintLine(1, "</Row>")

            For Row = 0 To objDataTable.Rows.Count - 1
                PrintLine(1, "<Row>")
                For Col = 1 To TotCol
                    If IsNumericFldCol(Col, NumericCols) = False Then
                        PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & objDataTable.Rows(Row).Item(Col - 1).ToString & "</Data></Cell>")
                    Else
                        PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""Number"">" & objDataTable.Rows(Row).Item(Col - 1).ToString & "</Data></Cell>")

                        'PrintLine(1, "<Cell><Data ss:Type=""String"" x:Ticked=""1"">" & objDataTable.Rows(Row).Item(Col - 1).ToString & "</Data></Cell>")

                        'PrintLine(1, "<Cell  ss:StyleID=""s21""><Data ss:Type=""DateTime"">" & objDataTable.Rows(Row).Item(Col - 1).ToString & "</Data></Cell>")
                    End If
                Next
                PrintLine(1, "</Row>")
            Next

            PrintLine(1, "</Table>")
            PrintLine(1, "</Worksheet>")
            PrintLine(1, "</Workbook>")

            FileClose(1)

            Dim lobjXls As New Excel.Application
            Dim lobjBook As Excel.Workbook
            lobjXls.DisplayAlerts = False
            lobjBook = lobjXls.Workbooks.Open(FileName)
            lobjBook.SaveAs(FileName, Excel.XlFileFormat.xlWorkbookNormal)
            lobjBook.Close()
            lobjBook = Nothing
            lobjXls = Nothing

            If DisplayFlag Then
                'Call ShellExecute(frmMain.Handle.ToInt32, "open", FileName, "", "0", SW_SHOWNORMAL)
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub PrintDGridPresentation(ByVal dgrid As DataGridView, ByVal FileName As String, Optional ByVal SheetName As String = "Report", Optional ByVal StoreasNum As String = "")

        Dim objDataTable As New Data.DataTable
        Dim TotCol As Integer, Col As Integer, Row As Integer
        Dim objStreamReader As StreamReader
        Dim lsTextContent As String
        Dim NumericCols() As String
        NumericCols = Split(StoreasNum, "|")

        Try

            objDataTable = dgrid.DataSource
            TotCol = objDataTable.Columns.Count
            If File.Exists(FileName) Then File.Delete(FileName)
            If File.Exists(FileName) = False Then

                FileOpen(1, FileName, OpenMode.Output)

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
            Else
                objStreamReader = File.OpenText(FileName)
                lsTextContent = objStreamReader.ReadToEnd()
                lsTextContent = Replace(lsTextContent.ToString, "</Workbook>", "")
                objStreamReader.Close()
                FileOpen(1, FileName, OpenMode.Output)
                PrintLine(1, lsTextContent.ToString)
            End If

            PrintLine(1, "<Worksheet ss:Name=""" & SheetName & """>")
            PrintLine(1, "<Table>")

            Dim liColIndex As Integer

            PrintLine(1, "<Row>")
            For Col = 1 To TotCol
                PrintLine(1, "<Cell ss:StyleID=""s1""><Data ss:Type=""String"">" & objDataTable.Columns(Col - 1).Caption.ToString & "</Data></Cell>")
                If objDataTable.Columns(Col - 1).Caption.ToString = "SLNO" Then liColIndex = Col - 1
            Next
            PrintLine(1, "</Row>")

            For Row = 0 To objDataTable.Rows.Count - 1
                PrintLine(1, "<Row>")
                For Col = 1 To TotCol
                    If Col - 1 = liColIndex Then
                        PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""Number"">" & (Row + 1) & "</Data></Cell>")
                    ElseIf IsNumericFldCol(Col, NumericCols) = False Then
                        If IsDate(objDataTable.Rows(Row).Item(Col - 1).ToString) Then
                            PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & Format(CDate(objDataTable.Rows(Row).Item(Col - 1).ToString), "dd-MMM-yyyy") & "</Data></Cell>")
                        Else
                            PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & objDataTable.Rows(Row).Item(Col - 1).ToString & "</Data></Cell>")
                        End If
                    Else
                        PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""Number"">" & objDataTable.Rows(Row).Item(Col - 1).ToString & "</Data></Cell>")
                    End If
                Next
                PrintLine(1, "</Row>")
            Next

            PrintLine(1, "</Table>")
            PrintLine(1, "</Worksheet>")
            PrintLine(1, "</Workbook>")

            FileClose(1)
            ' Call ShellExecute(frmMain.Handle.ToInt32, "open", FileName, "", "0", SW_SHOWNORMAL)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub SqlToXml_customized(ByVal SqlStr As String, ByVal ColFormat As String, ByVal ExcelFileName As String, ByVal SheetName As String, _
                              Optional ByVal ShowFlag As Boolean = True, _
                              Optional ByVal DelFlag As Boolean = True)
        Dim ds As New DataSet
        Dim objFile As File
        Dim objStreamReader As StreamReader
        Dim lsTextContent As String
        Dim lsColFormat() As String

        Dim n As Integer = 0, i As Integer = 0, j As Integer = 0, c As Integer = 0
        Dim lbBoldFlag As Boolean, lbPlainFlag As Boolean
        Dim lsTxt As String = ""

        Try

            lsColFormat = Split(ColFormat, ",")

            If objFile.Exists(ExcelFileName) = True And DelFlag = True Then objFile.Delete(ExcelFileName)

            If objFile.Exists(ExcelFileName) = False Then
                FileOpen(1, ExcelFileName, OpenMode.Append)

                PrintLine(1, "<?xml version=""1.0"" encoding=""utf-8""?>")
                PrintLine(1, "<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:html=""http://www.w3.org/TR/REC-html40"">")

                PrintLine(1, "<Styles>")
                PrintLine(1, "<Style ss:ID=""s1"">")
                PrintLine(1, "<Alignment ss:Horizontal=""Center"" ss:Vertical=""Bottom""/>")
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
                PrintLine(1, "<Style ss:ID=""s3"">")
                'PrintLine(1, "<Alignment ss:Horizontal=""left"" ss:Vertical=""centre""/>")
                PrintLine(1, "<Font x:Family=""Swiss"" ss:Bold=""1""/>")
                PrintLine(1, "</Style>")

                'PrintLine(1, "<Style ss:ID=""s21"">")
                'PrintLine(1, "<Borders>")
                'PrintLine(1, "<Border ss:PosiZtion=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
                'PrintLine(1, "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
                'PrintLine(1, "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
                'PrintLine(1, "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
                'PrintLine(1, "</Borders>")
                'PrintLine(1, "<NumberFormat ss:Format=""Short Date""/>")
                'PrintLine(1, "</Style>")

                PrintLine(1, "</Styles>")
            Else
                objStreamReader = File.OpenText(ExcelFileName)
                lsTextContent = objStreamReader.ReadToEnd()
                lsTextContent = lsTextContent.ToString.Replace("</Workbook>", "")
                objStreamReader.Close()
                FileOpen(1, ExcelFileName, OpenMode.Output)
                PrintLine(1, lsTextContent.ToString)
            End If

            PrintLine(1, "<Worksheet ss:Name=""" & SheetName & """>")

            ds = gfDataSet(SqlStr, "data", gOdbcConn)

            With ds.Tables("data")
                PrintLine(1, "<Table ss:ExpandedColumnCount=" & Chr(34) & .Columns.Count + 2 & Chr(34) & " " & _
                             "ss:ExpandedRowCount=" & Chr(34) & .Rows.Count + n + 2 & Chr(34) & ">")

                PrintLine(1, "<Row>")
                'PrintLine(1, "<Cell ss:StyleID=""s1""><Data ss:Type=""String"" x:Ticked=""1"">SNo</Data></Cell>")

                For i = 0 To .Columns.Count - 1
                    PrintLine(1, "<Cell ss:StyleID=""s1""><Data ss:Type=""String"" x:Ticked=""1"">" & .Columns(i).ColumnName & "</Data></Cell>")
                Next i

                PrintLine(1, "</Row>")

                For i = 0 To .Rows.Count - 1
                    PrintLine(1, "<Row>")
                    'PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & i + 1 & "</Data></Cell>")

                    For j = 0 To .Columns.Count - 1
                        lsTxt = .Rows(i).Item(j).ToString
                        lsTxt = lsTxt.Replace("<", "< ")
                        lsTxt = lsTxt.Replace(">", " >")

                        Select Case lsColFormat(j)
                            Case "N"
                                PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""Number"">" & lsTxt & "</Data></Cell>")
                            Case "S"
                                PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & lsTxt & "</Data></Cell>")
                            Case "E"
                                PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"">" & lsTxt & "</Data></Cell>")
                            Case "D"
                                PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"">" & lsTxt & "</Data></Cell>")
                        End Select

                    Next j

                    PrintLine(1, "</Row>")
                Next i
            End With

            PrintLine(1, "</Table>")
            PrintLine(1, "</Worksheet>")
            PrintLine(1, "</Workbook>")

            FileClose(1)
            FileClose(2)

            'If ShowFlag = True Then ShellExecute(frmMain.Handle.ToInt32, "open", ExcelFileName, "", "0", SW_SHOWNORMAL)
        Catch ex As Exception
            FileClose(1)
            ' objMail.GF_Mail(gsMailFrom, gsMailTo, gsMailSub, ex.Message, gsFormName, "SqlToXml")
            MsgBox(ex.Message)
        End Try
    End Sub


    Public Sub SqlToXml(ByVal SqlStr As String, ByVal ExcelFileName As String, ByVal SheetName As String, _
                              Optional ByVal ShowFlag As Boolean = True, _
                              Optional ByVal DelFlag As Boolean = True)
        Dim ds As New DataSet
        Dim objFile As File
        Dim objStreamReader As StreamReader
        Dim lsTextContent As String

        Dim n As Integer = 0, i As Integer = 0, j As Integer = 0, c As Integer = 0
        Dim lbBoldFlag As Boolean, lbPlainFlag As Boolean
        Dim lsTxt As String = ""

        Try
            If objFile.Exists(ExcelFileName) = True And DelFlag = True Then objFile.Delete(ExcelFileName)

            If objFile.Exists(ExcelFileName) = False Then
                FileOpen(1, ExcelFileName, OpenMode.Append)

                PrintLine(1, "<?xml version=""1.0"" encoding=""utf-8""?>")
                PrintLine(1, "<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:html=""http://www.w3.org/TR/REC-html40"">")

                PrintLine(1, "<Styles>")
                PrintLine(1, "<Style ss:ID=""s1"">")
                PrintLine(1, "<Alignment ss:Horizontal=""Center"" ss:Vertical=""Bottom""/>")
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
                PrintLine(1, "<Style ss:ID=""s3"">")
                'PrintLine(1, "<Alignment ss:Horizontal=""left"" ss:Vertical=""centre""/>")
                PrintLine(1, "<Font x:Family=""Swiss"" ss:Bold=""1""/>")
                PrintLine(1, "</Style>")
                PrintLine(1, "</Styles>")
            Else
                objStreamReader = File.OpenText(ExcelFileName)
                lsTextContent = objStreamReader.ReadToEnd()
                lsTextContent = lsTextContent.ToString.Replace("</Workbook>", "")
                objStreamReader.Close()
                FileOpen(1, ExcelFileName, OpenMode.Output)
                PrintLine(1, lsTextContent.ToString)
            End If

            PrintLine(1, "<Worksheet ss:Name=""" & SheetName & """>")

            ds = gfDataSet(SqlStr, "data", gOdbcConn)

            With ds.Tables("data")
                PrintLine(1, "<Table ss:ExpandedColumnCount=" & Chr(34) & .Columns.Count + 2 & Chr(34) & " " & _
                             "ss:ExpandedRowCount=" & Chr(34) & .Rows.Count + n + 2 & Chr(34) & ">")

                PrintLine(1, "<Row>")
                PrintLine(1, "<Cell ss:StyleID=""s1""><Data ss:Type=""String"" x:Ticked=""1"">SNo</Data></Cell>")

                For i = 0 To .Columns.Count - 1
                    PrintLine(1, "<Cell ss:StyleID=""s1""><Data ss:Type=""String"" x:Ticked=""1"">" & .Columns(i).ColumnName & "</Data></Cell>")
                Next i

                PrintLine(1, "</Row>")

                For i = 0 To .Rows.Count - 1
                    PrintLine(1, "<Row>")
                    PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & i + 1 & "</Data></Cell>")

                    For j = 0 To .Columns.Count - 1
                        lsTxt = .Rows(i).Item(j).ToString
                        lsTxt = lsTxt.Replace("<", "< ")
                        lsTxt = lsTxt.Replace(">", " >")

                        PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & lsTxt & "</Data></Cell>")

                        'If IsNumericFldCol(Col, NumericCols) = False Then
                        '    PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & objDataTable.Rows(Row).Item(Col - 1).ToString & "</Data></Cell>")
                        'Else
                        '    PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""Number"">" & objDataTable.Rows(Row).Item(Col - 1).ToString & "</Data></Cell>")
                        'End If


                    Next j

                    PrintLine(1, "</Row>")
                Next i
            End With

            PrintLine(1, "</Table>")
            PrintLine(1, "</Worksheet>")
            PrintLine(1, "</Workbook>")

            FileClose(1)
            FileClose(2)

            'If ShowFlag = True Then ShellExecute(frmMain.Handle.ToInt32, "open", ExcelFileName, "", "0", SW_SHOWNORMAL)
        Catch ex As Exception
            FileClose(1)
            ' objMail.GF_Mail(gsMailFrom, gsMailTo, gsMailSub, ex.Message, gsFormName, "SqlToXml")
            MsgBox(ex.Message)
        End Try
    End Sub


    Function IsNumericFldCol(ByVal ColPosition As Integer, ByVal NumericCols() As String) As Boolean
        Dim Temp As Integer
        For Temp = 0 To UBound(NumericCols)
            If ColPosition = Val(NumericCols(Temp)) Then
                IsNumericFldCol = True
                Exit Function
            End If
        Next Temp
        IsNumericFldCol = False
    End Function

    Public Sub PrintDGViewXML(ByVal dgrid As DataGridView, ByVal FileName As String, Optional ByVal SheetName As String = "Report", Optional ByVal StoreasNum As String = "")
        Dim objDataTable As New Data.DataTable
        Dim TotCol As Integer, Col As Integer, Row As Integer
        Dim objStreamReader As StreamReader
        Dim lsTextContent As String
        Dim NumericCols() As String
        NumericCols = Split(StoreasNum, "|")

        Try

            objDataTable = dgrid.DataSource
            TotCol = objDataTable.Columns.Count
            If File.Exists(FileName) Then File.Delete(FileName)
            If File.Exists(FileName) = False Then

                FileOpen(1, FileName, OpenMode.Output)

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
            Else
                objStreamReader = File.OpenText(FileName)
                lsTextContent = objStreamReader.ReadToEnd()
                lsTextContent = Replace(lsTextContent.ToString, "</Workbook>", "")
                objStreamReader.Close()
                FileOpen(1, FileName, OpenMode.Output)
                PrintLine(1, lsTextContent.ToString)
            End If

            PrintLine(1, "<Worksheet ss:Name=""" & SheetName & """>")
            PrintLine(1, "<Table>")

            PrintLine(1, "<Row>")
            For Col = 1 To TotCol
                PrintLine(1, "<Cell ss:StyleID=""s1""><Data ss:Type=""String"">" & objDataTable.Columns(Col - 1).Caption.ToString & "</Data></Cell>")
            Next
            PrintLine(1, "</Row>")

            For Row = 0 To objDataTable.Rows.Count - 1
                PrintLine(1, "<Row>")
                For Col = 1 To TotCol
                    If IsNumericFldCol(Col, NumericCols) = False Then
                        PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & objDataTable.Rows(Row).Item(Col - 1).ToString & "</Data></Cell>")
                    Else
                        PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""Number"">" & objDataTable.Rows(Row).Item(Col - 1).ToString & "</Data></Cell>")
                    End If
                Next
                PrintLine(1, "</Row>")
            Next

            PrintLine(1, "</Table>")
            PrintLine(1, "</Worksheet>")
            PrintLine(1, "</Workbook>")

            FileClose(1)
            'Call ShellExecute(frmMain.Handle.ToInt32, "open", FileName, "", "0", SW_SHOWNORMAL)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    'Func For Exporting FlexGrid (Having merged Cell) To Excel using XML
    Private Class FlexCell
        Public CellFlag As String = ""
        Public MergeRow As Integer = 0
        Public MergeCol As Integer = 0
    End Class
    'Export FlexGrid To Excel using XML
    'Public Sub PrintFGridXMLMerge_old(ByVal dgrid As AxMSFlexGridLib.AxMSFlexGrid, ByVal FileName As String, Optional ByVal SheetName As String = "Report", Optional ByVal StoreasNum As String = "")
    '    Dim objDataTable As New Data.DataTable
    '    Dim TotCol As Integer, Col As Integer, Row As Integer
    '    Dim objStreamReader As StreamReader
    '    Dim lsTextContent As String

    '    Dim Header(,) As FlexCell
    '    Dim RowStart() As Integer

    '    Dim i As Integer, j As Integer
    '    Dim lnCurrCol As Integer, lnCurrRow As Integer, lsCurrFlag As String
    '    Dim lsValue As String
    '    Dim lsNxtValue As String
    '    Dim lsTxt As String = ""

    '    Dim NumericCols() As String
    '    NumericCols = Split(StoreasNum, "|")

    '    Try
    '        ReDim Header(dgrid.FixedRows, dgrid.Cols)
    '        ReDim RowStart(dgrid.FixedRows)

    '        'objDataTable = dgrid.DataSource
    '        TotCol = dgrid.Cols

    '        If File.Exists(FileName) = True Then File.Delete(FileName)

    '        If File.Exists(FileName) = False Then

    '            FileOpen(1, FileName, OpenMode.Output)

    '            PrintLine(1, "<?xml version=""1.0"" encoding=""utf-8""?>")
    '            PrintLine(1, "<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:html=""http://www.w3.org/TR/REC-html40"">")

    '            PrintLine(1, "<Styles>")
    '            PrintLine(1, "<Style ss:ID=""s1"">")
    '            PrintLine(1, "<Font x:Family=""Swiss"" ss:Bold=""1""/>")
    '            PrintLine(1, "<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/>")
    '            PrintLine(1, "<Borders>")
    '            PrintLine(1, "<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
    '            PrintLine(1, "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
    '            PrintLine(1, "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
    '            PrintLine(1, "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
    '            PrintLine(1, "</Borders>")
    '            PrintLine(1, "</Style>")
    '            PrintLine(1, "<Style ss:ID=""s2"">")
    '            PrintLine(1, "<Borders>")
    '            PrintLine(1, "<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
    '            PrintLine(1, "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
    '            PrintLine(1, "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
    '            PrintLine(1, "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
    '            PrintLine(1, "</Borders>")
    '            PrintLine(1, "</Style>")
    '            PrintLine(1, "</Styles>")
    '        Else
    '            objStreamReader = File.OpenText(FileName)
    '            lsTextContent = objStreamReader.ReadToEnd()
    '            lsTextContent = lsTextContent.ToString.Replace("</Workbook>", "")
    '            objStreamReader.Close()
    '            FileOpen(1, FileName, OpenMode.Output)
    '            PrintLine(1, lsTextContent.ToString)
    '        End If

    '        With dgrid
    '            For i = 0 To .FixedRows - 1
    '                RowStart(i) = 0

    '                For j = 0 To .Cols - 1
    '                    Header(i, j) = New FlexCell
    '                Next j
    '            Next i

    '            For i = 0 To .FixedRows - 1
    '                For j = 0 To .Cols - 1
    '                    lsValue = .get_TextMatrix(i, j)
    '                    lsCurrFlag = Header(i, j).CellFlag

    '                    If lsCurrFlag = "" Then
    '                        lnCurrRow = i
    '                        lnCurrCol = j
    '                        lsCurrFlag = "Y"

    '                        Header(i, j).CellFlag = lsCurrFlag
    '                        Header(i, j).MergeRow = 0
    '                        Header(i, j).MergeCol = 0

    '                        For Col = j + 1 To .Cols - 1
    '                            lsNxtValue = .get_TextMatrix(i, Col)

    '                            If lsValue <> lsNxtValue Or Header(i, Col).CellFlag <> "" Then
    '                                Exit For
    '                            Else
    '                                lnCurrCol += 1
    '                                Header(i, j).MergeCol = lnCurrCol.ToString

    '                                Header(i, Col).CellFlag = "N"
    '                                Header(i, Col).MergeRow = 0
    '                                Header(i, Col).MergeCol = 0
    '                            End If
    '                        Next Col

    '                        For Row = i + 1 To .FixedRows - 1
    '                            lsNxtValue = .get_TextMatrix(Row, j)

    '                            If lsValue <> lsNxtValue Then
    '                                Exit For
    '                            Else
    '                                lnCurrRow += 1
    '                                Header(i, j).MergeRow = lnCurrRow.ToString

    '                                Header(Row, j).CellFlag = "N"
    '                                Header(Row, j).MergeRow = 0
    '                                Header(Row, j).MergeCol = 0
    '                            End If
    '                        Next Row

    '                        Header(i, j).MergeRow = lnCurrRow - i
    '                        Header(i, j).MergeCol = lnCurrCol - j

    '                        For Row = i + 1 To lnCurrRow
    '                            For Col = j + 1 To lnCurrCol
    '                                Header(Row, Col).CellFlag = "N"
    '                                Header(Row, Col).MergeRow = 0
    '                                Header(Row, Col).MergeCol = 0
    '                            Next Col
    '                        Next Row

    '                        If j = 0 And i < .FixedRows - 1 Then RowStart(i + 1) = lnCurrCol
    '                    End If
    '                Next j
    '            Next i
    '        End With

    '        PrintLine(1, "<Worksheet ss:Name=""" & SheetName & """>")
    '        PrintLine(1, "<Table ss:ExpandedColumnCount=" & Chr(34) & dgrid.Cols + 1 & Chr(34) & " " & _
    '                     "ss:ExpandedRowCount=" & Chr(34) & dgrid.Rows + 1 & Chr(34) & ">")

    '        For Row = 0 To dgrid.FixedRows - 1
    '            PrintLine(1, "<Row>")

    '            For Col = 0 To dgrid.Cols - 1
    '                If Header(Row, Col).CellFlag = "Y" Then
    '                    lsTxt = ""
    '                    lsTxt &= "<Cell "

    '                    If Row <> 0 And Col <> 0 Then
    '                        lsTxt &= IIf(RowStart(Row) <> 0 Or Header(Row - 1, Col - 1).MergeRow <> 0, "ss:Index=" & Chr(34) & RowStart(Row) + 2 & Chr(34) & " ", "")
    '                    End If

    '                    lsTxt &= IIf(Header(Row, Col).MergeCol = 0, "", "ss:MergeAcross=" & Chr(34) & Header(Row, Col).MergeCol & Chr(34) & " ")
    '                    lsTxt &= IIf(Header(Row, Col).MergeRow = 0, "", "ss:MergeDown=" & Chr(34) & Header(Row, Col).MergeRow & Chr(34) & " ")
    '                    lsTxt &= "ss:StyleID=""s1"">"
    '                    lsTxt &= "<Data ss:Type=""String"">" & dgrid.get_TextMatrix(Row, Col) & "</Data></Cell>"

    '                    PrintLine(1, lsTxt)
    '                End If
    '            Next Col

    '            PrintLine(1, "</Row>")
    '        Next Row

    '        For Row = dgrid.FixedRows To dgrid.Rows - 1
    '            PrintLine(1, "<Row>")
    '            For Col = 0 To dgrid.Cols - 1
    '                If IsNumericFldCol(Col, NumericCols) = False Then
    '                    PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & dgrid.get_TextMatrix(Row, Col) & "</Data></Cell>")
    '                Else
    '                    PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"">" & dgrid.get_TextMatrix(Row, Col) & "</Data></Cell>")
    '                End If
    '            Next
    '            PrintLine(1, "</Row>")
    '        Next

    '        PrintLine(1, "</Table>")
    '        PrintLine(1, "</Worksheet>")
    '        PrintLine(1, "</Workbook>")

    '        FileClose(1)

    '        Call ShellExecute(frmMain.Handle.ToInt32, "open", FileName, "", "0", SW_SHOWNORMAL)
    '    Catch ex As Exception
    '        FileClose(1)
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub
    'Export FlexGrid To Excel using XML
    'Public Sub PrintFGridXMLMerge(ByVal dgrid As AxMSFlexGridLib.AxMSFlexGrid, ByVal FileName As String, ByVal SheetName As String, Optional ByVal DelFlag As Boolean = True, Optional ByVal ShowFlag As Boolean = True)
    '    Dim objDataTable As New Data.DataTable
    '    Dim TotCol As Integer, Col As Integer, Row As Integer
    '    'Dim objFile As File
    '    Dim objStreamReader As StreamReader
    '    Dim lsTextContent As String

    '    Dim Header(,) As FlexCell
    '    Dim RowStart(,) As Integer

    '    Dim i As Integer, j As Integer
    '    Dim lnCurrCol As Integer, lnCurrRow As Integer, lsCurrFlag As String
    '    Dim lsValue As String
    '    Dim lsNxtValue As String
    '    Dim lsTxt As String = ""

    '    Try
    '        ReDim Header(dgrid.FixedRows, dgrid.Cols)
    '        ReDim RowStart(dgrid.FixedRows, dgrid.Cols)

    '        'objDataTable = dgrid.DataSource
    '        TotCol = dgrid.Cols

    '        If File.Exists(FileName) = True And DelFlag = True Then File.Delete(FileName)

    '        If File.Exists(FileName) = False Then
    '            FileOpen(1, FileName, OpenMode.Output)

    '            PrintLine(1, "<?xml version=""1.0"" encoding=""utf-8""?>")
    '            PrintLine(1, "<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:html=""http://www.w3.org/TR/REC-html40"">")

    '            PrintLine(1, "<Styles>")
    '            PrintLine(1, "<Style ss:ID=""s1"">")
    '            PrintLine(1, "<Alignment ss:Horizontal=""Center"" ss:Vertical=""Bottom""/>")
    '            PrintLine(1, "<Font x:Family=""Swiss"" ss:Bold=""1""/>")
    '            PrintLine(1, "<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/>")
    '            PrintLine(1, "<Borders>")
    '            PrintLine(1, "<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
    '            PrintLine(1, "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
    '            PrintLine(1, "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
    '            PrintLine(1, "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
    '            PrintLine(1, "</Borders>")
    '            PrintLine(1, "</Style>")
    '            PrintLine(1, "<Style ss:ID=""s2"">")
    '            PrintLine(1, "<Borders>")
    '            PrintLine(1, "<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
    '            PrintLine(1, "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
    '            PrintLine(1, "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
    '            PrintLine(1, "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
    '            PrintLine(1, "</Borders>")
    '            PrintLine(1, "</Style>")
    '            PrintLine(1, "</Styles>")
    '        Else
    '            objStreamReader = File.OpenText(FileName)
    '            lsTextContent = objStreamReader.ReadToEnd()
    '            lsTextContent = lsTextContent.ToString.Replace("</Workbook>", "")
    '            objStreamReader.Close()
    '            FileOpen(1, FileName, OpenMode.Output)
    '            PrintLine(1, lsTextContent.ToString)
    '        End If

    '        With dgrid
    '            For i = 0 To .FixedRows - 1
    '                For j = 0 To .Cols - 1
    '                    RowStart(i, j) = 0
    '                    Header(i, j) = New FlexCell
    '                Next j
    '            Next i

    '            For i = 0 To .FixedRows - 1
    '                For j = 0 To .Cols - 1
    '                    lsValue = .get_TextMatrix(i, j)
    '                    lsCurrFlag = Header(i, j).CellFlag

    '                    If lsCurrFlag = "" Then
    '                        lnCurrRow = i
    '                        lnCurrCol = j
    '                        lsCurrFlag = "Y"

    '                        Header(i, j).CellFlag = lsCurrFlag
    '                        Header(i, j).MergeRow = 0
    '                        Header(i, j).MergeCol = 0

    '                        For Col = j + 1 To .Cols - 1
    '                            lsNxtValue = .get_TextMatrix(i, Col)

    '                            If lsValue <> lsNxtValue Or Header(i, Col).CellFlag <> "" Then
    '                                Exit For
    '                            Else
    '                                lnCurrCol += 1
    '                                Header(i, j).MergeCol = lnCurrCol.ToString

    '                                Header(i, Col).CellFlag = "N"
    '                                Header(i, Col).MergeRow = 0
    '                                Header(i, Col).MergeCol = 0
    '                            End If
    '                        Next Col

    '                        For Row = i + 1 To .FixedRows - 1
    '                            lsNxtValue = .get_TextMatrix(Row, j)

    '                            If lsValue <> lsNxtValue Then
    '                                Exit For
    '                            Else
    '                                lnCurrRow += 1
    '                                Header(i, j).MergeRow = lnCurrRow.ToString

    '                                Header(Row, j).CellFlag = "N"
    '                                Header(Row, j).MergeRow = 0
    '                                Header(Row, j).MergeCol = 0
    '                            End If
    '                        Next Row

    '                        Header(i, j).MergeRow = lnCurrRow - i
    '                        Header(i, j).MergeCol = lnCurrCol - j

    '                        For Row = i + 1 To lnCurrRow
    '                            If lnCurrCol <> 0 Then
    '                                If RowStart(Row, lnCurrCol - 1) <> 0 Then RowStart(Row, lnCurrCol - 1) = 0
    '                                If lnCurrCol < .Cols - 1 Then RowStart(Row, lnCurrCol + 1) = lnCurrCol + 1
    '                            Else
    '                                If lnCurrCol < .Cols - 1 Then RowStart(Row, lnCurrCol + 1) = 1
    '                            End If

    '                            For Col = j + 1 To lnCurrCol
    '                                Header(Row, Col).CellFlag = "N"
    '                                Header(Row, Col).MergeRow = 0
    '                                Header(Row, Col).MergeCol = 0
    '                            Next Col
    '                        Next Row

    '                        'If j = 0 And i < .FixedRows - 1 Then RowStart(i + 1) = lnCurrCol
    '                    End If
    '                Next j
    '            Next i
    '        End With

    '        PrintLine(1, "<Worksheet ss:Name=""" & SheetName & """>")
    '        PrintLine(1, "<Table ss:ExpandedColumnCount=" & Chr(34) & dgrid.Cols + 1 & Chr(34) & " " & _
    '                     "ss:ExpandedRowCount=" & Chr(34) & dgrid.Rows + 1 & Chr(34) & ">")

    '        For Row = 0 To dgrid.FixedRows - 1
    '            PrintLine(1, "<Row>")

    '            For Col = 0 To dgrid.Cols - 1
    '                If Header(Row, Col).CellFlag = "Y" Then
    '                    lsTxt = ""
    '                    lsTxt &= "<Cell "

    '                    If Row <> 0 And Col <> 0 Then
    '                        'lsTxt &= IIf(RowStart(Row, Col) <> 0 Or Header(Row - 1, Col - 1).MergeRow <> 0, "ss:Index=" & Chr(34) & RowStart(Row, Col) + 2 & Chr(34) & " ", "")
    '                        lsTxt &= IIf(RowStart(Row, Col) <> 0, "ss:Index=" & Chr(34) & RowStart(Row, Col) + 1 & Chr(34) & " ", "")
    '                    End If

    '                    lsTxt &= IIf(Header(Row, Col).MergeCol = 0, "", "ss:MergeAcross=" & Chr(34) & Header(Row, Col).MergeCol & Chr(34) & " ")
    '                    lsTxt &= IIf(Header(Row, Col).MergeRow = 0, "", "ss:MergeDown=" & Chr(34) & Header(Row, Col).MergeRow & Chr(34) & " ")
    '                    lsTxt &= "ss:StyleID=""s1"">"
    '                    lsTxt &= "<Data ss:Type=""String"">" & dgrid.get_TextMatrix(Row, Col) & "</Data></Cell>"

    '                    PrintLine(1, lsTxt)
    '                End If
    '            Next Col

    '            PrintLine(1, "</Row>")
    '        Next Row

    '        For Row = dgrid.FixedRows To dgrid.Rows - 1
    '            PrintLine(1, "<Row>")
    '            For Col = 0 To dgrid.Cols - 1
    '                If dgrid.get_ColAlignment(Col) <> 7 Or IsNumeric(dgrid.get_TextMatrix(Row, Col)) = False Then
    '                    PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & dgrid.get_TextMatrix(Row, Col) & "</Data></Cell>")
    '                ElseIf IsNumeric(dgrid.get_TextMatrix(Row, Col)) = True Then
    '                    PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""Number"">" & dgrid.get_TextMatrix(Row, Col) & "</Data></Cell>")
    '                Else
    '                    PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & dgrid.get_TextMatrix(Row, Col) & "</Data></Cell>")
    '                End If
    '                'Else
    '                'PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"">" & dgrid.get_TextMatrix(Row, Col) & "</Data></Cell>")
    '                'End If
    '            Next
    '            PrintLine(1, "</Row>")
    '        Next

    '        PrintLine(1, "</Table>")
    '        PrintLine(1, "</Worksheet>")
    '        PrintLine(1, "</Workbook>")

    '        FileClose(1)

    '        If ShowFlag = True Then Call ShellExecute(frmMain.Handle.ToInt32, "open", FileName, "", "0", SW_SHOWNORMAL)
    '    Catch ex As Exception
    '        FileClose(1)
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub
    'Export DataGrid To Excel using XML

    Public Sub PrintMultiSheetExcelFormat(ByRef FileName As String, ByRef ShtCount As Short, ByRef sheetname As Object, ByRef Sql() As Object, Optional ByRef PBar As Boolean = False, Optional ByRef BarName As Object = Nothing)
        Dim fs As New Scripting.FileSystemObject
        Dim objXls As New Excel.Application
        Dim objBook As Excel.Workbook
        Dim objShe As Excel.Worksheet
        Dim lRow, lCol As Short
        Dim lbAddFlag As Boolean
        Dim lShtNo As Short
        Dim lFit As Short
        Dim lobjDT As New DataTable
        Dim PBARVAL As Integer

        If Not fs.FileExists(FileName) Then
            objBook = objXls.Workbooks.Add
            lbAddFlag = True
        Else
            If MessageBox.Show("The Path you have mentioned Already contains a file named " & FileName & Chr(13) & "Would you Like to Replace the Existing File?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = MsgBoxResult.No Then
                Exit Sub
            End If
            fs.DeleteFile(FileName)
            objBook = objXls.Workbooks.Add
            lbAddFlag = True
            objBook.SaveAs(FileName, 1)
            lbAddFlag = False
            objBook.Save()
        End If

        For lShtNo = 0 To ShtCount - 1
            If lShtNo > 2 Then
                objShe = objXls.Worksheets.Add
            Else
                objShe = objXls.ActiveWorkbook.Worksheets("sheet" & lShtNo + 1)
            End If

            objShe.Name = sheetname(lShtNo)
            lobjDT = gobjConnection.GetDataTable(Sql(lShtNo))


            PBARVAL = 0
            BarName = 0
            If PBar Then If lobjDT.Rows.Count > 0 Then BarName.maximum = lobjDT.Rows.Count
            With lobjDT
                lRow = 0
                Dim lindex As Integer
                lindex = 1
                For lRow = 0 To lobjDT.Rows.Count - 1

                    For lCol = 0 To lobjDT.Columns.Count - 1
                        If lRow = 0 Then
                            objShe.Cells._Default(1, lCol + 1) = .Columns(lCol).ColumnName
                            objShe.Cells._Default(1, lCol + 1).Font.Bold = True
                        End If

                    Next

                    For lCol = 0 To lobjDT.Columns.Count - 1
                        If IsDate(lobjDT.Rows(lRow).Item(lCol).ToString) Then
                            If Mid(.Rows(lRow).Item(lCol).ToString, 3, 1) = "-" And Mid(.Rows(lRow).Item(lCol).ToString, 6, 1) = "-" Then
                                objShe.Cells._Default(lindex + 1, lCol + 1) = "'" & Format(CDate(.Rows(lRow).Item(lCol)), "dd-MM-yyyy")
                            End If
                        Else
                            objShe.Cells._Default(lindex + 1, lCol + 1) = "'" & .Rows(lRow).Item(lCol).ToString
                        End If
                        If lRow > 0 Then
                            objShe.Cells._Default(lindex, lCol + 1).Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
                            objShe.Cells._Default(lindex + 1, lCol + 1).Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black)
                        End If
                    Next lCol
                    If PBar Then
                        If BarName.Value < BarName.Maximum Then
                            BarName.Value = PBARVAL
                            PBARVAL = PBARVAL + 1
                        End If
                    End If
                    lindex = lindex + 1
                Next lRow
                For lFit = 1 To .Columns.Count
                    objShe.Columns._Default(lFit).AutoFit()
                    If .Rows.Count > 0 Then
                        objShe.Cells._Default(1, lFit).Interior.ColorIndex = 6
                        objShe.Cells._Default(1, lFit).Font.ColorIndex = 3
                    End If
                Next lFit
            End With

        Next lShtNo

        If Not lbAddFlag Then
            objBook.Save()
        Else
            objBook.SaveAs(FileName, 1)
        End If

        objBook.Save()
        objXls.Workbooks.Close()
        objXls.Quit()
        objShe = Nothing
        objBook = Nothing
        objXls = Nothing
    End Sub

    Public Sub PrintDGridXML(ByVal dgrid As DataGrid, ByVal FileName As String, Optional ByVal SheetName As String = "Report", Optional ByVal StoreasNum As String = "")
        Dim objDataTable As New Data.DataTable
        Dim TotCol As Integer, Col As Integer, Row As Integer
        Dim objStreamReader As StreamReader
        Dim lsTextContent As String
        Dim NumericCols() As String
        NumericCols = Split(StoreasNum, "|")

        Try

            objDataTable = dgrid.DataSource
            TotCol = objDataTable.Columns.Count
            If File.Exists(FileName) Then File.Delete(FileName)
            If File.Exists(FileName) = False Then

                FileOpen(1, FileName, OpenMode.Output)

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
            Else
                objStreamReader = File.OpenText(FileName)
                lsTextContent = objStreamReader.ReadToEnd()
                lsTextContent = Replace(lsTextContent.ToString, "</Workbook>", "")
                objStreamReader.Close()
                FileOpen(1, FileName, OpenMode.Output)
                PrintLine(1, lsTextContent.ToString)
            End If

            PrintLine(1, "<Worksheet ss:Name=""" & SheetName & """>")
            PrintLine(1, "<Table>")

            PrintLine(1, "<Row>")
            For Col = 1 To TotCol
                PrintLine(1, "<Cell ss:StyleID=""s1""><Data ss:Type=""String"">" & objDataTable.Columns(Col - 1).Caption.ToString & "</Data></Cell>")
            Next
            PrintLine(1, "</Row>")

            For Row = 0 To objDataTable.Rows.Count - 1
                PrintLine(1, "<Row>")
                For Col = 1 To TotCol
                    If IsNumericFldCol(Col, NumericCols) = False Then
                        PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & objDataTable.Rows(Row).Item(Col - 1).ToString & "</Data></Cell>")
                    Else
                        PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""Number"">" & objDataTable.Rows(Row).Item(Col - 1).ToString & "</Data></Cell>")
                    End If
                Next
                PrintLine(1, "</Row>")
            Next

            PrintLine(1, "</Table>")
            PrintLine(1, "</Worksheet>")
            PrintLine(1, "</Workbook>")

            FileClose(1)
            'Call ShellExecute(frmMain.Handle.ToInt32, "open", FileName, "", "0", SW_SHOWNORMAL)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Public Sub PrintFGridXMLMerge(ByVal dgrid As AxMSFlexGridLib.AxMSFlexGrid, ByVal FileName As String, Optional ByVal SheetName As String = "Report", Optional ByVal DisplayFlag As Boolean = True, Optional ByVal DeleteFlag As Boolean = True)
    '    'Declarations
    '    Dim objDataTable As New Data.DataTable
    '    Dim TotCol As Integer, Col As Integer, Row As Integer
    '    Dim objStreamReader As StreamReader
    '    Dim lsTextContent As String

    '    Dim Header(,) As FlexCell
    '    Dim RowStart(,) As Integer

    '    Dim i As Integer, j As Integer
    '    Dim lnCurrCol As Integer, lnCurrRow As Integer, lsCurrFlag As String
    '    Dim lsValue As String
    '    Dim lsNxtValue As String
    '    Dim lsTxt As String = ""

    '    Try

    '        ReDim Header(dgrid.FixedRows, dgrid.Cols)
    '        ReDim RowStart(dgrid.FixedRows, dgrid.Cols)

    '        'objDataTable = dgrid.DataSource
    '        TotCol = dgrid.Cols

    '        If DeleteFlag Then
    '            If File.Exists(FileName) = True Then File.Delete(FileName)
    '        End If

    '        If File.Exists(FileName) = False Then

    '            FileOpen(1, FileName, OpenMode.Output)

    '            PrintLine(1, "<?xml version=""1.0"" encoding=""utf-8""?>")
    '            PrintLine(1, "<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:html=""http://www.w3.org/TR/REC-html40"">")

    '            PrintLine(1, "<Styles>")
    '            PrintLine(1, "<Style ss:ID=""s1"">")
    '            PrintLine(1, "<Alignment ss:Horizontal=""Center"" ss:Vertical=""Bottom""/>")
    '            PrintLine(1, "<Font x:Family=""Swiss"" ss:Bold=""1""/>")
    '            PrintLine(1, "<Interior ss:Color=""#C0C0C0"" ss:Pattern=""Solid""/>")
    '            PrintLine(1, "<Borders>")
    '            PrintLine(1, "<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
    '            PrintLine(1, "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
    '            PrintLine(1, "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
    '            PrintLine(1, "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
    '            PrintLine(1, "</Borders>")
    '            PrintLine(1, "</Style>")
    '            PrintLine(1, "<Style ss:ID=""s2"">")
    '            PrintLine(1, "<Borders>")
    '            PrintLine(1, "<Border ss:Position=""Bottom"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
    '            PrintLine(1, "<Border ss:Position=""Left"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
    '            PrintLine(1, "<Border ss:Position=""Right"" ss:LineStyle=""Continuous"" ss:Weight=""1"" />")
    '            PrintLine(1, "<Border ss:Position=""Top"" ss:LineStyle=""Continuous"" ss:Weight=""1"" /> ")
    '            PrintLine(1, "</Borders>")
    '            PrintLine(1, "</Style>")
    '            PrintLine(1, "</Styles>")
    '        Else
    '            objStreamReader = File.OpenText(FileName)
    '            lsTextContent = objStreamReader.ReadToEnd()
    '            lsTextContent = lsTextContent.ToString.Replace("</Workbook>", "")
    '            objStreamReader.Close()
    '            FileOpen(1, FileName, OpenMode.Output)
    '            PrintLine(1, lsTextContent.ToString)
    '        End If

    '        With dgrid
    '            For i = 0 To .FixedRows - 1
    '                For j = 0 To .Cols - 1
    '                    RowStart(i, j) = 0
    '                    Header(i, j) = New FlexCell
    '                Next j
    '            Next i

    '            For i = 0 To .FixedRows - 1
    '                For j = 0 To .Cols - 1
    '                    lsValue = .get_TextMatrix(i, j)
    '                    lsCurrFlag = Header(i, j).CellFlag

    '                    If lsCurrFlag = "" Then
    '                        lnCurrRow = i
    '                        lnCurrCol = j
    '                        lsCurrFlag = "Y"

    '                        Header(i, j).CellFlag = lsCurrFlag
    '                        Header(i, j).MergeRow = 0
    '                        Header(i, j).MergeCol = 0

    '                        For Col = j + 1 To .Cols - 1
    '                            lsNxtValue = .get_TextMatrix(i, Col)

    '                            If lsValue <> lsNxtValue Or Header(i, Col).CellFlag <> "" Then
    '                                Exit For
    '                            Else
    '                                lnCurrCol += 1
    '                                Header(i, j).MergeCol = lnCurrCol.ToString

    '                                Header(i, Col).CellFlag = "N"
    '                                Header(i, Col).MergeRow = 0
    '                                Header(i, Col).MergeCol = 0
    '                            End If
    '                        Next Col

    '                        For Row = i + 1 To .FixedRows - 1
    '                            lsNxtValue = .get_TextMatrix(Row, j)

    '                            If lsValue <> lsNxtValue Then
    '                                Exit For
    '                            Else
    '                                lnCurrRow += 1
    '                                Header(i, j).MergeRow = lnCurrRow.ToString

    '                                Header(Row, j).CellFlag = "N"
    '                                Header(Row, j).MergeRow = 0
    '                                Header(Row, j).MergeCol = 0
    '                            End If
    '                        Next Row

    '                        Header(i, j).MergeRow = lnCurrRow - i
    '                        Header(i, j).MergeCol = lnCurrCol - j

    '                        For Row = i + 1 To lnCurrRow
    '                            If lnCurrCol <> 0 Then
    '                                If RowStart(Row, lnCurrCol - 1) <> 0 Then RowStart(Row, lnCurrCol - 1) = 0
    '                                If lnCurrCol < .Cols - 1 Then RowStart(Row, lnCurrCol + 1) = lnCurrCol + 1
    '                            Else
    '                                If lnCurrCol < .Cols - 1 Then RowStart(Row, lnCurrCol + 1) = 1
    '                            End If

    '                            For Col = j + 1 To lnCurrCol
    '                                Header(Row, Col).CellFlag = "N"
    '                                Header(Row, Col).MergeRow = 0
    '                                Header(Row, Col).MergeCol = 0
    '                            Next Col
    '                        Next Row

    '                        'If j = 0 And i < .FixedRows - 1 Then RowStart(i + 1) = lnCurrCol
    '                    End If
    '                Next j
    '            Next i
    '        End With

    '        PrintLine(1, "<Worksheet ss:Name=""" & SheetName & """>")
    '        PrintLine(1, "<Table ss:ExpandedColumnCount=" & Chr(34) & dgrid.Cols + 1 & Chr(34) & " " & _
    '                     "ss:ExpandedRowCount=" & Chr(34) & dgrid.Rows + 1 & Chr(34) & ">")

    '        For Row = 0 To dgrid.FixedRows - 1
    '            PrintLine(1, "<Row>")

    '            For Col = 0 To dgrid.Cols - 1
    '                If Header(Row, Col).CellFlag = "Y" Then
    '                    lsTxt = ""
    '                    lsTxt &= "<Cell "

    '                    If Row <> 0 And Col <> 0 Then
    '                        'lsTxt &= IIf(RowStart(Row, Col) <> 0 Or Header(Row - 1, Col - 1).MergeRow <> 0, "ss:Index=" & Chr(34) & RowStart(Row, Col) + 2 & Chr(34) & " ", "")
    '                        lsTxt &= IIf(RowStart(Row, Col) <> 0, "ss:Index=" & Chr(34) & RowStart(Row, Col) + 1 & Chr(34) & " ", "")
    '                    End If

    '                    lsTxt &= IIf(Header(Row, Col).MergeCol = 0, "", "ss:MergeAcross=" & Chr(34) & Header(Row, Col).MergeCol & Chr(34) & " ")
    '                    lsTxt &= IIf(Header(Row, Col).MergeRow = 0, "", "ss:MergeDown=" & Chr(34) & Header(Row, Col).MergeRow & Chr(34) & " ")
    '                    lsTxt &= "ss:StyleID=""s1"">"
    '                    lsTxt &= "<Data ss:Type=""String"">" & dgrid.get_TextMatrix(Row, Col) & "</Data></Cell>"

    '                    PrintLine(1, lsTxt)
    '                End If
    '            Next Col

    '            PrintLine(1, "</Row>")
    '        Next Row

    '        For Row = dgrid.FixedRows To dgrid.Rows - 1
    '            PrintLine(1, "<Row>")
    '            For Col = 0 To dgrid.Cols - 1
    '                If dgrid.get_ColAlignment(Col) <> 7 Or IsNumeric(dgrid.get_TextMatrix(Row, Col)) = False Then
    '                    PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & dgrid.get_TextMatrix(Row, Col) & "</Data></Cell>")
    '                Else
    '                    PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""Number"">" & dgrid.get_TextMatrix(Row, Col) & "</Data></Cell>")
    '                End If
    '                'Else
    '                'PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"">" & dgrid.get_TextMatrix(Row, Col) & "</Data></Cell>")
    '                'End If
    '            Next
    '            PrintLine(1, "</Row>")
    '        Next

    '        PrintLine(1, "</Table>")
    '        PrintLine(1, "</Worksheet>")
    '        PrintLine(1, "</Workbook>")

    '        FileClose(1)

    '        If DisplayFlag Then
    '            Call ShellExecute(frmMain.Handle.ToInt32, "open", FileName, "", "0", SW_SHOWNORMAL)
    '        End If

    '    Catch ex As Exception
    '        FileClose(1)
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub

End Module