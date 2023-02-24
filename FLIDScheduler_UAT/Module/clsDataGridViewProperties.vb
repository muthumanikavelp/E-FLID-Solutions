Public Class clsDataGridViewProperties
    Public gs_ColumnWidth As String
    Public gs_ColumnAlignment As String
    Public gs_Qry As String
    Public gs_Table As String
    Public Function FillGridView(ByRef gs_DataGrid As DataGridView)
        FillGridView = Nothing
        Dim fobjdataset As New DataSet
        Dim fstotColumnWidth As String()
        Dim fsColumnAlignment As String()
        Dim fiTotCount As Integer

        gs_DataGrid.DataSource = Nothing
        fobjdataset = gfDataSet(gs_Qry, gs_Table, gOdbcConn)
        gs_DataGrid.DataSource = fobjdataset.Tables(gs_Table)

        fstotColumnWidth = Split(gs_ColumnWidth, "|")
        fsColumnAlignment = Split(gs_ColumnAlignment, "|")

        If UBound(fstotColumnWidth) = UBound(fsColumnAlignment) Then
            For fiTotCount = 0 To UBound(fstotColumnWidth)
                gs_DataGrid.Columns(fiTotCount).HeaderText = fobjdataset.Tables(gs_Table).Columns(fiTotCount).ColumnName
                gs_DataGrid.Columns(fiTotCount).Width = fstotColumnWidth(fiTotCount)
                gs_DataGrid.Columns(fiTotCount).ReadOnly = True

                If Trim(fsColumnAlignment(fiTotCount).ToString) = 1 Then
                    gs_DataGrid.Columns(fiTotCount).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                ElseIf Trim(fsColumnAlignment(fiTotCount).ToString) = 2 Then
                    gs_DataGrid.Columns(fiTotCount).DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopCenter
                ElseIf Trim(fsColumnAlignment(fiTotCount).ToString) = 3 Then
                    gs_DataGrid.Columns(fiTotCount).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                End If
                gs_DataGrid.ColumnHeadersDefaultCellStyle.NullValue = ""
            Next
        End If
    End Function
End Class
