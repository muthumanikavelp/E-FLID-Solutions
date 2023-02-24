Imports System.IO
Imports System.IO.FileStream
Imports System.Data.Odbc
Imports System.Data
Imports System.Data.OleDb
Module modGlobalODBC
#Region "Global Declaration"


    Public ServerDetails As String
    Public ServerDetails1 As String
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

    Public gProjectName As String = "KIT"
    Public softcode As String = "KIT"
    Public gUid As Integer
    Public gUserFullName As String = ""
    Public gUserRights As String
    Public GFormName As String
    Public gsPrintFilepath As String = ""

    Public gOdbcConn As New OdbcConnection
    Public gOdbcConn1 As New OdbcConnection
    Public gOdbcConnBIZ As New OdbcConnection
    Public gOdbcDAdp As New OdbcDataAdapter
    Public gOdbcCmd As New OdbcCommand
    Public gOdbcCmdBIZ As New OdbcCommand

    Public gFso As New FileIO.FileSystem

    Public gsReportPath As String = "C:\Execute\"
    Public txt As Long
    Public gsPacketStatus As String
    Public Sqlstr As String
    Public lsCond As String

    Public gbVerification As Boolean
    Public gbAddFlag As Boolean
    Public gbMChqFlag As Boolean
    Public gbEditFlag As Boolean
    Public lni As Integer
    Public glPacketCheckList As Long
    Public glAuditCheckList As Long
    Public lsPayMode As String
    Public gsPacketNo As String
    Public gsEnteredBy As String
    Public gsEntryDtFrom As String
    Public gsEntryDtTo As String
    Public gsSpoolMonth As String
    Public gstitle As String = "KIT"
    Public gsWorkCode As String
    Public lsUserid As String = gobjSecurity.LoginUserCode

    Dim RCount As Integer
    Dim empid As Integer
    Dim empname As String

    Public gn_pre_month As String
    Public gs_database As String
    Public gs_changemonth As String
    Public gs_CycleMonth As String

    Public gnEvenColor As Long = RGB(220, 200, 100)
    Public gnOddColor As Long = RGB(175, 210, 175)

    'Public gnEvenColor As Long = RGB(220, 200, 500)
    'Public gnOddColor As Long = RGB(175, 210, 675)

    Public Const gnAuth As Integer = 1
    Public Const gnReject As Integer = 2

    'Mail Information'

    Public RA_From As String = "ra.loansppu@bizconceptsindia.com"
    Public Sys_To As String = "sysadmin@gnsaindia.com"
    Public Mail_Subject As String = "Error Message"

#End Region
    'For calling the Main form
    '' ''Public Sub Main()
    '' ''    Call ConOpenOdbc(ServerDetails)
    '' ''    Call ConOpenOdbcBIZ(ServerDetailsBIZ)
    '' ''    Try

    '' ''        Dim Security As New GNSASecurity.clsForm

    '' ''        'If UCase(Application.LocalUserAppDataPath) <> UCase("C:\EXEC\QA") Then
    '' ''        '    MsgBox("File path Error!" & Chr(13) & _
    '' ''        '     "File Only from C:\EXEC\QA can access", vbInformation, gProjectName)
    '' ''        '    Exit Sub
    '' ''        'End If

    '' ''        Security.clsDbIP = DbIP             ' commented for testing
    '' ''        Security.clsDbUID = DbUId           ' commented for testing
    '' ''        Security.clsDbPWD = DbPwd           ' commented for testing
    '' ''        Security.clsDbName = DbName         ' commented for testing

    '' ''        If Command() <> "" Then Security.clsEmpId = Val(Command)

    '' ''        Security.LoadLogin()                   ' commented for testing

    '' ''        gUserName = UCase(Security.clsEmpShortName)
    '' ''        gUserRights = UCase(Security.ModuleRights("RECEIPT"))
    '' ''        gUId = UCase(Security.clsEmpName)
    '' ''        If gUserName <> "" Then
    '' ''frmMain.ShowDialog()
    '' ''        End If
    '' ''    Catch ex As Exception
    '' ''        objMail.GF_Mail(RA_From, Sys_To, Mail_Subject, ex.Message, GFormName, "Query")
    '' ''        MsgBox(ex.Message)
    '' ''    End Try
    '' ''End Sub
    'To open the Connection
    Public Sub ConOpenOdbc(ByVal ServerDetails As String)
        If gOdbcConn.State = ConnectionState.Closed Then
            gOdbcConn.ConnectionString = ServerDetails
            gOdbcConn.Open()
            gOdbcCmd.Connection = gOdbcConn
        End If
        'empid = Security.clsEmpId
    End Sub
    'To open the Connection
    Public Sub ConOpenOdbcBIZ(ByVal ServerDetailsBIZ As String)
        If gOdbcConnBIZ.State = ConnectionState.Closed Then
            gOdbcConnBIZ.ConnectionString = ServerDetailsBIZ
            gOdbcConnBIZ.Open()
            gOdbcCmdBIZ.Connection = gOdbcConnBIZ
        End If
        'empid = Security.clsEmpId
    End Sub
    'To Close the Connection
    Public Sub ConCloseOdbc(ByVal ServerDetails As String)
        If gOdbcConn.State = ConnectionState.Open Then
            gOdbcConn.Close()
        End If
    End Sub
    'To Execute Query and return as datareader
    Public Function gfExecuteQry(ByVal strsql As String, ByVal odbcConn As OdbcConnection)
        Dim objCommand As OdbcCommand
        Dim objDataReader As OdbcDataReader
        objCommand = New OdbcCommand(strsql, odbcConn)
        Try
            objDataReader = objCommand.ExecuteReader()
            objCommand.Dispose()
            Return objDataReader
        Catch ex As Exception
            MsgBox(ex.Message)
            Return (0)
        End Try

    End Function
    'To Execute Query and return value as boolean
    Public Function gfExecuteQryBln(ByVal strsql As String, ByVal odbcConn As OdbcConnection) As Boolean
        gOdbcCmd = New OdbcCommand(strsql, odbcConn)
        Dim objDataReader As OdbcDataReader
        Try
            objDataReader = gOdbcCmd.ExecuteReader()
            If objDataReader.HasRows Then
                gfExecuteQryBln = True
            Else
                gfExecuteQryBln = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Function
        End Try
    End Function
    'To Execute Query and return value as string
    Public Function gfExecuteScalar(ByVal strsql As String, ByVal odbcConn As OdbcConnection) As String
        Dim StrVal As String
        Dim objCommand As OdbcCommand
        objCommand = New OdbcCommand(strsql, odbcConn)

        Try
            If IsDBNull(objCommand.ExecuteScalar()) Or IsNothing(objCommand.ExecuteScalar()) Then
                StrVal = ""
            Else
                StrVal = objCommand.ExecuteScalar()
            End If

            objCommand.Dispose()
            Return StrVal

        Catch ex As Exception
            MsgBox(ex.Message)
            Return 0
        End Try

    End Function

    Public Sub LoadXLSheet(ByVal FileName As String, ByVal objCbo As ComboBox)

        'Dim objXL As New Excel.Application

        'Dim i As Integer

        'objCbo.Items.Clear()
        'objXL.Workbooks.Open(FileName)

        'For i = 1 To objXL.ActiveWorkbook.Worksheets.Count
        '    objCbo.Items.Add(objXL.ActiveWorkbook.Worksheets(i).name)
        'Next i

        'objXL.Workbooks.Close()

        'GC.Collect()
        'GC.WaitForPendingFinalizers()
        'objXL.Quit()
        'System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objXL)
        'objXL = Nothing

    End Sub

    'To Execute Query and return value as integer
    Public Function gfInsertQry(ByVal strsql As String, ByVal odbcConn As OdbcConnection) As Integer
        Dim recAffected As Long
        gOdbcCmd = New OdbcCommand(strsql, odbcConn)
        gOdbcCmd.CommandType = CommandType.Text
        Try
            recAffected = gOdbcCmd.ExecuteNonQuery()
            Return recAffected
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Function
        End Try
    End Function
    'To Bind values to Datagrid
    Public Sub gpPopGrid(ByVal GridName As DataGrid, ByVal Qry As String, ByVal odbcConn As OdbcConnection)
        Dim lobjDataTable As New DataTable
        Dim lobjDataView As New DataView
        Dim lobjDataSet As New DataSet
        Dim lobjDataAdapter As New Odbc.OdbcDataAdapter
        Try
            lobjDataAdapter = New OdbcDataAdapter(Qry, odbcConn)
            lobjDataSet = New DataSet("TBL")
            lobjDataAdapter.Fill(lobjDataSet, "TBL")
            lobjDataTable = lobjDataSet.Tables(0)
            lobjDataView = New DataView(lobjDataTable)
            GridName.DataSource = lobjDataView
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    'To Bind values to Datagrid
    Public Sub gpPopGridView(ByVal GridName As DataGridView, ByVal Qry As String, ByVal odbcConn As OdbcConnection)
        Dim lda As New Odbc.OdbcDataAdapter(Qry, odbcConn)
        Dim lds As New DataSet
        Dim ldt As DataTable
        Try
            lda.Fill(lds, "tbl")
            ldt = lds.Tables("tbl")
            GridName.DataSource = ldt
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    'To filter single quote in the give text

    'To Clear control in a form
    Public Sub frmCtrClear(ByVal frmName As Form)
        Dim ctrl As Control
        For Each ctrl In frmName.Controls
            If ctrl.Tag <> "*" Then
                If TypeOf ctrl Is TextBox Then ctrl.Text = ""
                If TypeOf ctrl Is ComboBox Then
                    ctrl.Text = ""
                End If
                'If TypeOf ctrl Is CheckBox Then
                '    ctrl = False
                'End If 
            End If
        Next
    End Sub
    'To get Dataset
    Public Function gfDataSet(ByVal SQL As String, ByVal tblName As String, ByVal odbcConn As Odbc.OdbcConnection) As DataSet
        Dim objDataAdapter As New OdbcDataAdapter(SQL, odbcConn)
        Dim objDataSet As New DataSet
        Try
            objDataAdapter.Fill(objDataSet, tblName)
            Return objDataSet
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function
    'Binding combo
    Public Sub gpBindCombo(ByVal SQL As String, ByVal Dispfld As String, _
                               ByVal Valfld As String, ByRef ComboName As ComboBox, _
                                ByVal odbcConn As Odbc.OdbcConnection)

        Dim objDataAdapter As New OdbcDataAdapter
        Dim objCommand As New OdbcCommand
        Dim objDataTable As New Data.DataTable
        Try
            objCommand.Connection = odbcConn
            objCommand.CommandType = CommandType.Text
            objCommand.CommandText = SQL
            objDataAdapter.SelectCommand = objCommand
            objDataAdapter.Fill(objDataTable)
            ComboName.DataSource = objDataTable
            ComboName.DisplayMember = Dispfld
            ComboName.ValueMember = Valfld
            ComboName.SelectedIndex = -1

        Catch ex As Exception
            MsgBox(ex.Message)
            objDataTable.Dispose()
            objCommand.Dispose()
            objDataAdapter.Dispose()
        End Try
    End Sub
    'Validating for Integer only
    Public Function gfIntEntryOnly(ByVal e As System.Windows.Forms.KeyPressEventArgs) As Boolean
        Select Case Asc(e.KeyChar)
            Case 48 To 57, 8, 22
                gfIntEntryOnly = False
            Case Else
                gfIntEntryOnly = True
        End Select
    End Function
    Public Function gfAmountEntryOnly(ByVal e As System.Windows.Forms.KeyPressEventArgs, ByVal txt As TextBox) As Boolean
        Select Case Asc(e.KeyChar)
            Case 48 To 57, 8, 46
                If Asc(e.KeyChar) = 46 Then
                    If InStr(txt.Text, ".") = 0 Then
                        gfAmountEntryOnly = False
                    Else
                        gfAmountEntryOnly = True
                    End If
                Else
                    gfAmountEntryOnly = False
                End If
            Case Else
                gfAmountEntryOnly = True
        End Select
    End Function
    'To Get the DataTable
    Public Function GetDataTable(ByVal SqlQry As String) As DataTable
        Dim lobjDataTable As New DataTable
        Dim lobjDataView As New DataView
        Dim lobjDataSet As New DataSet
        Dim lobjDataAdapter As New Odbc.OdbcDataAdapter
        GetDataTable = Nothing
        Try

            gOdbcDAdp = New OdbcDataAdapter(SqlQry, gOdbcConn)
            lobjDataSet = New DataSet("TBL")
            gOdbcDAdp.Fill(lobjDataSet, "TBL")
            lobjDataTable = lobjDataSet.Tables(0)
            lobjDataView = New DataView(lobjDataTable)
            Return lobjDataTable

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    'Disables the addition of rows in the given DataGrid
    Public Sub DisableAddNew(ByRef dg As DataGrid, _
                                    ByRef Frm As Form)
        ' Disable addnew capability on the grid.
        ' Note that AllowEdit and AllowDelete can be disabled
        ' by adding or changing the "AllowNew" property to
        ' AllowDelete or AllowEdit.
        Dim cm As CurrencyManager = _
           CType(Frm.BindingContext(dg.DataSource, dg.DataMember), _
                 CurrencyManager)
        CType(cm.List, DataView).AllowNew = False
    End Sub
    ' Aligns the given text in specified format
    Public Function AlignTxt(ByVal txt As String, ByVal Length As Integer, ByVal Alignment As Integer) As String
        Dim X As String = ""

        Select Case Alignment
            Case 1
                Return LSet(txt, Length)
            Case 4
                Return CSet(txt, Length)
            Case 7
                Return RSet(txt, Length)
            Case Else
                Return (0)
        End Select
    End Function
    ' Center Align the Given Text
    Public Function CSet(ByVal txt As String, ByVal PaperChrWidth As Integer) As String
        Dim s As String                 ' Temporary String Variable
        Dim l As Integer                ' Length of the String
        If Len(txt) > PaperChrWidth Then
            CSet = Left(txt, PaperChrWidth)
        Else
            l = (PaperChrWidth - Len(txt)) / 2
            s = RSet(txt, l + Len(txt))
            CSet = Space(PaperChrWidth - Len(s))
            CSet = s + CSet
        End If
    End Function
    Public Function SwapChkSum(ByVal txt As String) As Double
        Dim TempTxt As String
        Dim TempChkSum As Double
        Dim i As Long

        TempTxt = txt
        TempChkSum = 0

        For i = 1 To Len(TempTxt)
            TempChkSum = TempChkSum + Asc(Mid(TempTxt, i, 1)) + (i - 1)
        Next i

        SwapChkSum = TempChkSum
    End Function
    Public Function SwapChkSumNew(ByVal txt As String) As Double
        Dim TempTxt As String
        Dim TempChkSum As Double
        Dim i As Long

        TempTxt = txt
        TempChkSum = 0

        For i = 1 To Len(TempTxt)
            TempChkSum = TempChkSum + Asc(Mid(TempTxt, i, 1)) + (i)
        Next i

        SwapChkSumNew = TempChkSum
    End Function
    Public Function ConvUcase(ByVal keychar As String) As String
        Select Case keychar
            Case "a" To "z"
                ConvUcase = keychar.ToUpper
            Case Else
                ConvUcase = keychar
        End Select
    End Function
    Public Function gfAgeing(ByVal FromDt As Date, ByVal ToDt As Date) As Long
        Dim m As Long, n As Long

        n = Val(gfExecuteScalar("select count(*) from prf_mst_tholiday " _
            & "where holiday_date >= '" & Format(FromDt, "yyyy-MM-dd") & "' " _
            & "and holiday_date <= '" & Format(ToDt, "yyyy-MM-dd") & "' " _
            & "and delete_flag is null", gOdbcConn))

        m = DateDiff("d", FromDt, ToDt)

        gfAgeing = m - n
    End Function
    Public Sub Kill_Excel()
        Dim proc As System.Diagnostics.Process
        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
            proc.Kill()
        Next
    End Sub
    'For Aging Calculation
    Public Function Ageing(ByVal FromDt As Date, ByVal ToDt As Date) As Long
        Dim m As Long, n As Long

        n = Val(gfExecuteScalar("select count(*) from prf_mst_tholiday " _
            & "where holiday_date >= '" & Format(FromDt, "yyyy-MM-dd") & "' " _
            & "and holiday_date <= '" & Format(ToDt, "yyyy-MM-dd") & "' " _
            & "and delete_flag is null", gOdbcConn))

        m = DateDiff("d", FromDt, ToDt)

        Ageing = m - n
    End Function
    'For Aging Date Calculation
    Public Function AgeingDt(ByVal dt As Date, ByVal Interval As Long) As Date
        Dim m As Long, N As Long, i As Long
        Dim mdToDate As Date
        i = 0

        Do
            mdToDate = DateAdd("d", ((Interval + i) * -1), dt)

            N = Val(gfExecuteScalar("select count(*) from prf_mst_tholiday " _
                & "where holiday_date <= '" & Format(dt, "yyyy-MM-dd") & "' " _
                & "and holiday_date >= '" & Format(mdToDate, "yyyy-MM-dd") & "' " _
                & "and delete_flag is null", gOdbcConn))


            m = DateDiff("d", mdToDate, dt) - N

            'i = i + 1
            i = i + (Interval - m)
        Loop Until m = Interval

        AgeingDt = DateAdd("d", -(Interval + i), dt)
    End Function
    'For Getting Loan Type 
    Public Function gfLoanType(ByVal lsLoanNo As String)
        Dim lsLoanType As String = ""
        Dim lnLoanLen As Integer

        lnLoanLen = lsLoanNo.Length
        If IsNumeric(lsLoanNo) Then
            lsLoanType = "H"
        Else
            lsLoanType = Mid(lsLoanNo, 1, 1)
            If lsLoanType = "A" Then
                lsLoanType = "A"
            ElseIf lsLoanType = "C" Or lsLoanType = "D" Then
                lsLoanType = "C"
            Else
                lsLoanType = "P"
            End If

        End If
        Return lsLoanType
    End Function
    'For Zip A File
    Public Sub gp_WinZip(ByVal password As String, ByVal DirPath As String, ByVal ZipPath As String)
        Dim FileName As String
        Dim X As String
        Dim lb_Flag As Boolean
        Try
            Const ZIPEXE = "C:\Program Files\WinZip\WINZIP32.EXE "

            'DirPath = txtAttachment1.Text
            'ZipPath = "c:\WinZip"

            'Password = "Citibank" & Mid(Format(Now, "dd-MM-yyyy"), 1, 2)
            If Dir(ZipPath, vbDirectory) = "" Then
                MkDir(ZipPath)
            End If
            FileName = Dir(DirPath, vbNormal)
            While FileName <> ""
                X = ZIPEXE & " -a -s" & password & " " & ZipPath & "\" & Mid(FileName, 1, Len(FileName) - 4) & ".zip " & FileName
                Shell(X)
                FileName = Dir()
                lb_Flag = True
            End While

            If lb_Flag = True Then
                MsgBox("Successfully Created at " & ZipPath, MsgBoxStyle.Information, gProjectName)
            Else
                MsgBox("ZIP Procedd Faild", MsgBoxStyle.Information, gProjectName)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Excel To DS :Created Date :23-02-2009 :Created By :Ilaya
    Public Function gpExcelDataset(ByVal Qry As String, ByVal Excelpath As String) As DataTable
        Dim fOleDbConString As String = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" & Excelpath & ";" + "Extended Properties=Excel 8.0;"
        Dim lobjDataTable As New DataTable
        Dim lobjDataSet As New DataSet
        Dim lobjDataAdapter As New OleDbDataAdapter

        lobjDataAdapter = New OleDbDataAdapter(Qry, fOleDbConString)
        lobjDataSet = New DataSet("TBL")
        lobjDataAdapter.Fill(lobjDataSet, "TBL")
        lobjDataTable = lobjDataSet.Tables(0)
        Return lobjDataTable

    End Function
    'LoadExcelSheet
    Public Sub gfLoadXLSheet(ByVal FileName As String, ByVal objCbo As ComboBox)
        ''Dim objXL As New Excel.Application
        'Dim i As Integer

        'objCbo.Items.Clear()
        'objXL.Workbooks.Open(FileName)

        'For i = 1 To objXL.ActiveWorkbook.Worksheets.Count
        '    objCbo.Items.Add(objXL.ActiveWorkbook.Worksheets(i).name)
        'Next i

        'objXL.Workbooks.Close()

        'GC.Collect()
        'GC.WaitForPendingFinalizers()
        'objXL.Quit()
        'System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objXL)
        'objXL = Nothing
    End Sub

    'AutoFillCombo :Created Date :24-02-2009 :Created By :Ilaya
    Public Sub gpAutoFillCombo(ByVal cboBox As ComboBox)

        Dim lnLenght As Long

        With cboBox

            lnLenght = .Text.Length

            .SelectedIndex = .FindString(.Text)

            .SelectionStart = lnLenght

            .SelectionLength = Math.Abs(.Text.Length - lnLenght)

        End With

    End Sub
    'AutoFillCombo :Created Date :24-02-2009 Created By :Ilaya
    Public Sub gpAutoFindCombo(ByVal cboBox As ComboBox)
        cboBox.SelectedIndex = cboBox.FindString(cboBox.Text)
    End Sub

    'Public Sub RowColor(ByVal ctrl As AxMSFlexGridLib.AxMSFlexGrid, ByVal StartRow As Integer, ByVal EndRow As Integer, ByVal BkColor As Long)
    '    Dim i As Integer, j As Integer

    '    Try
    '        With ctrl
    '            For i = StartRow To EndRow
    '                .Row = i

    '                For j = .FixedCols To .Cols - 1
    '                    .Col = j
    '                    .CellBackColor = ColorTranslator.FromWin32(BkColor)
    '                Next j
    '            Next i
    '        End With
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, gProjectName)
    '    End Try
    'End Sub

    Public Function gfAmtEntryOnly(ByVal e As System.Windows.Forms.KeyPressEventArgs) As Boolean
        Select Case Asc(e.KeyChar)
            Case 48 To 57, 8, 46
                gfAmtEntryOnly = False
            Case Else
                gfAmtEntryOnly = True
        End Select
    End Function

    'Binding combo
    Public Sub gpBindDGridCombo(ByVal SQL As String, ByVal Dispfld As String, _
                               ByVal Valfld As String, ByRef ComboName As DataGridViewComboBoxColumn, _
                                ByVal odbcConn As Odbc.OdbcConnection)

        Dim objDataAdapter As New OdbcDataAdapter
        Dim objCommand As New OdbcCommand
        Dim objDataTable As New Data.DataTable
        Try
            objCommand.Connection = odbcConn
            objCommand.CommandType = CommandType.Text
            objCommand.CommandText = SQL
            objDataAdapter.SelectCommand = objCommand
            objDataAdapter.Fill(objDataTable)
            ComboName.DataSource = objDataTable
            ComboName.DisplayMember = Dispfld
            ComboName.ValueMember = Valfld
            'ComboName.SelectedIndex = -1

        Catch ex As Exception
            MsgBox(ex.Message)
            objDataTable.Dispose()
            objCommand.Dispose()
            objDataAdapter.Dispose()
        End Try
    End Sub
    'DBF To DS :Created Date :03-03-2009 :Created By :Kali
    Public Function gpDBFDataset(ByVal Qry As String, ByVal DBFPath As String) As DataTable

        Dim fOleDbConString As String = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" & DBFPath & ";" + "Extended Properties=dBASE IV;User ID=Admin;Password=;"
        Dim lobjDataTable As New DataTable
        Dim lobjDataSet As New DataSet
        Dim lobjDataAdapter As New OleDbDataAdapter

        lobjDataAdapter = New OleDbDataAdapter(Qry, fOleDbConString)
        lobjDataSet = New DataSet("TBL")
        lobjDataAdapter.Fill(lobjDataSet, "TBL")
        lobjDataTable = lobjDataSet.Tables(0)
        Return lobjDataTable
    End Function
    'Chq Number validate :Created Date :23-02-2009 :Created By :Ilaya
    Public Function gfValidate_chqDate(ByVal chqdt As Date, ByVal CycleDate As Date) As Boolean
        Dim monthdt As Date = Nothing
        monthdt = DateAdd(DateInterval.Month, -6, CycleDate) 'Date less than 6 month from cycle date
        If IsDate(chqdt) Then
            If CDate(monthdt) <= CDate(chqdt) Then
                If CDate(chqdt) <= CDate(CycleDate) Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Public Function ffGetSlNo(ByVal BatchNo As Long, ByVal BusinessGid As Long) As Long
        Dim ds As DataSet = Nothing
        Dim llSlNo As Long = 0
        Dim lsSql As String = ""

        lsSql = ""
        lsSql &= " select max(chq_slno) from pre_trn_tbounce"
        lsSql &= " where batch_no = " & BatchNo & ""
        lsSql &= " and business_gid = " & BusinessGid & ""
        lsSql &= " and delete_flag is null"

        ds = gfDataSet(lsSql, "tblSlNo", gOdbcConn)
        If ds.Tables("tblSlNo").Rows.Count > 0 Then
            If Not ds.Tables("tblSlNo").Rows(0).Item(0).ToString = "" Then
                llSlNo = Val(ds.Tables("tblSlNo").Rows(0).Item(0).ToString) + 1
            Else
                llSlNo = "1"
            End If
        End If

        Return llSlNo
    End Function
End Module