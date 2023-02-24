Imports System.IO

Module modMIG
    Public Sub PrintDGridviewXML(ByVal dt As DataTable, ByVal FileName As String)
        Dim liTotCol As Integer, liCol As Integer, lRow As Long
        Dim lsSheetName As String
        Dim liSubSheet As Integer
        Dim lsOutputFile As String
        Dim lTotRow As Long

        liSubSheet = 0

        Try
            liTotCol = dt.Columns.Count
            lsSheetName = "Report"
            lsOutputFile = FileName & ".xls"
            If File.Exists(lsOutputFile) Then File.Delete(lsOutputFile)
            FileOpen(1, lsOutputFile, OpenMode.Output)
            PrintLine(1, "<?xml version=""1.0"" encoding=""utf-8""?>")
            PrintLine(1, "<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" xmlns:html=""http://www.w3.org/TR/REC-html40"">")
            PrintLine(1, "<Styles>")
            PrintLine(1, "<Style ss:ID=""s1"">")
            PrintLine(1, "<Font x:Family=""Swiss"" ss:Color=""#FF0000"" ss:Bold=""1""/>")
            PrintLine(1, "<Interior ss:Color=""#FFFF00"" ss:Pattern=""Solid""/>")
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
                PrintLine(1, "<Cell ss:StyleID=""s1""><Data ss:Type=""String"">" & dt.Columns(liCol - 1).Caption.ToString & "</Data></Cell>")
            Next
            PrintLine(1, "</Row>")
            For lRow = 0 To dt.Rows.Count - 1
                If lTotRow > 65000 Then
                    PrintLine(1, "</Table>")
                    PrintLine(1, "</Worksheet>")
                    PrintLine(1, "</Workbook>")
                    FileClose(1)
                    lTotRow = 1
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
                        PrintLine(1, "<Cell ss:StyleID=""s1""><Data ss:Type=""String"">" & dt.Columns(liCol - 1).Caption.ToString & "</Data></Cell>")
                    Next
                    PrintLine(1, "</Row>")
                End If
                PrintLine(1, "<Row>")
                For liCol = 1 To liTotCol
                    PrintLine(1, "<Cell ss:StyleID=""s2""><Data ss:Type=""String"" x:Ticked=""1"">" & dt.Rows(lRow).Item(liCol - 1).ToString & "</Data></Cell>")
                Next
                PrintLine(1, "</Row>")
                lTotRow = lTotRow + 1
            Next
            PrintLine(1, "</Table>")
            PrintLine(1, "</Worksheet>")
            PrintLine(1, "</Workbook>")
            FileClose(1)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Public Function QuoteFilter(ByVal txt As String) As String
        QuoteFilter = Trim(Replace(Replace(Replace(txt.ToString, "'", "''"), """", """"""), "\", "/"))
    End Function
End Module
