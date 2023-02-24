Imports System.Data.SqlClient
Imports System.IO

Public Class frmExcelImporter

    Dim loDBConnection As New iODBCconnection

    Dim gsTitle As String = "MIG"
    Dim fsColumnHeaders As String

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        'cmbSheet.Items.Clear()
        txtFileName.Text = ""
        txtStatus.Text = ""
        cboMIGMode.SelectedIndex = -1
        cboMIGMode.Focus()
    End Sub

    'Private Sub FetchSheetNamesFromXLFile()
    '    Try

    '        Dim lobjXls As New Excel.Application
    '        Dim lobjBook As Excel.Workbook
    '        Dim liIndex As Integer

    '        If txtFileName.Text <> "" Then
    '            cmbSheet.Items.Clear()

    '            lobjBook = lobjXls.Workbooks.Open(txtFileName.Text)

    '            For liIndex = 1 To lobjXls.ActiveWorkbook.Worksheets.Count
    '                cmbSheet.Items.Add(lobjXls.ActiveWorkbook.Worksheets(liIndex).Name)
    '            Next liIndex

    '            lobjXls.Workbooks.Close()
    '            lobjXls.Quit()
    '        End If

    '        lobjBook = Nothing
    '        lobjXls = Nothing

    '        GC.Collect()
    '        GC.WaitForPendingFinalizers()

    '        cmbSheet.Focus()

    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message, gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try

    'End Sub

    Private Sub btnGetFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetFile.Click
        With OpenFileDialog1
            .Filter = "Excel Files|*.*"
            .Title = "Select Files to Import"
            .FileName = ""

            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                If .FileName <> "" Then txtFileName.Text = .FileName
            End If
            'If txtFileName.Text.Trim <> "" Then
            '    Call FetchSheetNamesFromXLFile()
            'End If
        End With
        cmbSheet.SelectedIndex = 0
    End Sub

    Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImport.Click

        If cboMIGMode.SelectedIndex = -1 Then
            MessageBox.Show("MIG MODE Can't be empty...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            cboMIGMode.Focus()
            Exit Sub
        End If

        If txtFileName.Text.Trim = "" Then
            MessageBox.Show("Please Select the Source File", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            btnGetFile.Focus()
            Exit Sub
        End If

        If cmbSheet.SelectedIndex = -1 Then
            MessageBox.Show("Please Select Sheet ", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            cmbSheet.Focus()
            Exit Sub
        End If

        txtStatus.Text = " Validating dump is in progress..."
        Application.DoEvents()

        cboMIGMode.Enabled = False
        btnGetFile.Enabled = False
        'cmbSheet.Enabled = False
        btnImport.Enabled = False

        Me.Cursor = Cursors.AppStarting

        If cboMIGMode.SelectedIndex = 0 Then                                                ' CBF 
            fsColumnHeaders = ""
            fsColumnHeaders &= ",CBFNo,Detail_Sno,CBF_OBF_Flag,Start_Date,End_Date,Project_Owner,Branch,Mode,Approval_Type,Is_Budgeted,Deviation_Amount"
            fsColumnHeaders &= ",CBF_Amount,Description,Raiser,Request_For,Remarks,Is_Branch_Single,Budget_Owner,PAR_PR_Description,Product_Service"
            fsColumnHeaders &= ",CBF_Details_Description,UOM,QTY,Unit_Price,Total_Amount,CBF_Details_Remarks,Chart_Of_Acc,FCCC,Budget_Line,Vendor,Product_Group,"

            CBF_Data_Migration()

        ElseIf cboMIGMode.SelectedIndex = 1 Then                                            ' PO 
            fsColumnHeaders = ""
            fsColumnHeaders &= ",SlNO,poheader_pono,podetails_cbfheader_cbfno,CBFDetail_Sno,poheader_date,poheader_enddate,poheader_raisor,poheader_projectmanager"
            fsColumnHeaders &= ",poheader_requestfor,poheader_ittype,poheader_vendor,poheader_vendor_note,poheader_over_total,poheader_type,poheader_termcond"
            fsColumnHeaders &= ",poheader_add_termandcond,podetails_prodservice,podetails_desc,podetails_uom,podetails_qty,podetails_unitprice,podetails_discount"
            fsColumnHeaders &= ",podetails_base_amt,podetails_tax1,podetails_tax2,podetails_tax3,podetails_total,Status,"

            PO_Data_Migration()

        ElseIf cboMIGMode.SelectedIndex = 2 Then                                            ' PO - SHIPMENT  
            fsColumnHeaders = ""
            fsColumnHeaders &= ",SlNO,poheader_pono,poshipment_shipmenttype,poshipment_branch,poshipment_remarks,podetails_qty,poshipment_incharge,"

            PO_Shipment_Data_Migration()

        ElseIf cboMIGMode.SelectedIndex = 3 Then                                            ' WO 
            fsColumnHeaders = ""
            fsColumnHeaders &= ",PO Number,CBF_header_CBFNo,CBFDetail_Sno,Date,Raisor,RequestFor,IT Type,Vendor Name,Vendor_note,Header Total"
            fsColumnHeaders &= ",Frequency Type,From_month,To_month,Type,Term_and_condition,Additional_term_and_condition,Product_service"
            fsColumnHeaders &= ",Description,Service_month,Percentage,Details_total,PO Detail Qty,Legacy Branch,"

            WO_Data_Migration()

        ElseIf cboMIGMode.SelectedIndex = 4 Then                                            ' ECF
            fsColumnHeaders = ""
            fsColumnHeaders &= ",Rowid,br_code,AUTH_DT,SF_BATCHNO,SF_DOCNUMB,SF_NUMBBIL,SF_BILNUMB,SF_VENCODE,SF_VENNAME,sf_actual_vendor,SF_BILDATE"
            fsColumnHeaders &= ",SF_BILSTDT,SF_BILEDDT,SF_BAMOUNT,SF_BILDISC,SF_NETAMNT,SF_RVENCODE,SF_RVENNAME,SF_VENDGRP,SF_STAT,EMPNAME,GN_ADDCODE"
            fsColumnHeaders &= ",sf_Status,sf_staxable,sf_servtax,sf_staxper,sf_paid,xpuBillNo,dedupBillNo,sertaxno,pan_no,CSTno,LSTno,WCTno,VATno"
            fsColumnHeaders &= ",sf_expmonth,provision_flag,vendorbranch_gid,SF_APPROVER_ID"

            'ECF_Data_Migration()


        ElseIf cboMIGMode.SelectedIndex = 5 Then                                            ' RETENTON
            fsColumnHeaders = ""
            fsColumnHeaders &= "SNO,INWARD DATE,REFERENCE NO,AMOUNT,PAID AMOUNT,BALANCE AMOUNT,RETENTION GL,VENDOR CODE,VENDOR NAME,INVOICE NO,BRANCH CODE"
            fsColumnHeaders &= ",MODULE CODE,NARRATION,DEPARTMENT,AGEING,Bucket,Remark"

            'ECF_Data_Migration()

        End If


        btnImport.Enabled = True
        'cmbSheet.Enabled = True
        btnGetFile.Enabled = True
        cboMIGMode.Enabled = True

        Me.Cursor = Cursors.Default

        Call btnClear_Click(Nothing, Nothing)

    End Sub


    Public Sub RETENTON_Data_Migration()
        Try
            Dim lsSQL As String = ""
            Dim lsFileName As String = ""
            Dim lsResult As String = ""
            Dim liDumpGid As Integer = 0
            Dim lobjOleDbConnection As New OleDb.OleDbConnection
            Dim lobjDataAdapter As OleDb.OleDbDataAdapter
            Dim lobjDataTable As DataTable
            Dim lobjDataSet As New DataSet
            Dim liError As Integer
            Dim liValid As Integer
            Dim liBranchGID As Integer

            Dim lsErrMessage As String = ""
            Dim lsErrNote As String = ""



            lsFileName = Mid(txtFileName.Text.Trim, InStrRev(txtFileName.Text.Trim, "\") + 1)

            Try
                If lobjOleDbConnection.State = 1 Then lobjOleDbConnection.Close()
                With lobjOleDbConnection

                    If Microsoft.VisualBasic.Right(lsFileName, 4) = "xlsx" Then
                        'read a 2007 file   
                        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                                             txtFileName.Text.Trim & ";" + ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
                    Else
                        'read a 97-2003 file   
                        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" & _
                                                txtFileName.Text.Trim & ";" + "Extended Properties=Excel 8.0;"
                    End If

                    .Open()

                End With

            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                lobjOleDbConnection.Close()
                Exit Sub
            End Try

            lsSQL = "SELECT * FROM [" & cmbSheet.Text & "$]"

            lobjDataAdapter = New OleDb.OleDbDataAdapter(lsSQL, lobjOleDbConnection)
            lobjDataAdapter.Fill(lobjDataSet, "DATA")
            lobjDataTable = lobjDataSet.Tables("DATA")
            lobjOleDbConnection.Close()

            Dim liIndex As Integer

            For liIndex = 0 To lobjDataTable.Columns.Count - 1
                If Not InStr(fsColumnHeaders.ToLower.Trim, "," & lobjDataTable.Columns(liIndex).ColumnName.ToString.Trim.ToLower & ",") > 0 Then
                    MessageBox.Show("WO Dump is not in the correct format...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If
            Next

            Dim lsSNO As String
            Dim lsINWARDDATE As String
            Dim lsREFERENCENO As String
            Dim lsAMOUNT As String
            Dim lsPAIDAMOUNT As String
            Dim lsBALANCEAMOUNT As String
            Dim lsRETENTIONGL As String
            Dim lsVENDORCODE As String
            Dim lsVENDORNAME As String
            Dim lsINVOICENO As String
            Dim lsBRANCHCODE As String
            Dim lsMODULECODE As String
            Dim lsNARRATION As String
            Dim lsDEPARTMENT As String
            Dim lsAGEING As String
            Dim lsBUCKET As String
            Dim lsRemark As String

            Dim liVendorGID As Integer
            Dim liRaiserGID As Integer
            Dim liECFGID As Integer
            Dim liInvoiceGID As Integer
            Dim liDebitLineGID As Integer
            Dim liCreditLineGID As Integer
            Dim liRetentionGID As Integer

            Dim lobjErrorDatatable As New DataTable

            'SNO	INWARD DATE	REFERENCE NO	AMOUNT	PAID AMOUNT	BALANCE AMOUNT	RETENTION GL	VENDOR CODE	VENDOR NAME	
            'INVOICE NO	BRANCH CODE	MODULE CODE	NARRATION	DEPARTMENT	AGEING	Bucket	Remark

            With lobjErrorDatatable
                .Columns.Add("SNO")
                .Columns.Add("INWARD DATE")
                .Columns.Add("REFERENCE NO")
                .Columns.Add("AMOUNT")
                .Columns.Add("PAID AMOUNT")
                .Columns.Add("BALANCE AMOUNT")
                .Columns.Add("RETENTION GL")
                .Columns.Add("VENDOR CODE")
                .Columns.Add("VENDOR NAME")
                .Columns.Add("INVOICE NO")
                .Columns.Add("BRANCH CODE")
                .Columns.Add("MODULE CODE")
                .Columns.Add("NARRATION")
                .Columns.Add("DEPARTMENT")
                .Columns.Add("AGEING")
                .Columns.Add("Bucket")
                .Columns.Add("Remark")

                .Columns.Add("Error")
                .Columns.Add("Error Note")
            End With

            For i As Integer = 0 To lobjDataTable.Rows.Count - 1
                lsErrMessage = ""
                lsErrNote = ""

                With lobjDataTable.Rows(i)
                    lsSNO = QuoteFilter(.Item("SNO").ToString)
                    lsINWARDDATE = QuoteFilter(.Item("INWARD DATE").ToString)
                    lsREFERENCENO = QuoteFilter(.Item("REFERENCE NO").ToString)
                    lsAMOUNT = QuoteFilter(.Item("AMOUNT").ToString)
                    lsPAIDAMOUNT = QuoteFilter(.Item("PAID AMOUNT").ToString)
                    lsBALANCEAMOUNT = QuoteFilter(.Item("BALANCE AMOUNT").ToString)
                    lsRETENTIONGL = QuoteFilter(.Item("RETENTION GL").ToString)
                    lsVENDORCODE = QuoteFilter(.Item("VENDOR CODE").ToString)
                    lsVENDORNAME = QuoteFilter(.Item("VENDOR NAME").ToString)
                    lsINVOICENO = QuoteFilter(.Item("INVOICE NO").ToString)
                    lsBRANCHCODE = QuoteFilter(.Item("BRANCH CODE").ToString)
                    lsMODULECODE = QuoteFilter(.Item("MODULE CODE").ToString)
                    lsNARRATION = QuoteFilter(.Item("NARRATION").ToString)
                    lsDEPARTMENT = QuoteFilter(.Item("DEPARTMENT").ToString)
                    lsAGEING = QuoteFilter(.Item("AGEING").ToString)
                    lsBUCKET = QuoteFilter(.Item("Bucket").ToString)
                    lsRemark = QuoteFilter(.Item("Remark").ToString)
                End With

                If lsINWARDDATE <> "" Then
                    If IsDate(lsINWARDDATE) Then
                        lsINWARDDATE = "'" & Format(CDate(lsINWARDDATE), "yyyy-MM-dd") & "'"
                    Else
                        lsINWARDDATE = "NULL"
                    End If
                Else
                    lsINWARDDATE = "NULL"
                End If

                ' SUPPLIER GID 
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL = " SELECT supplierheader_gid FROM asms_trn_tsupplierheader "
                lsSQL &= " WHERE supplierheader_name='" & lsVendorName & "' "  'Mid(lsVendorName, 1, InStr(lsVendorName, "-") - 1).Trim

                liVendorGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


                liRaiserGID = 1375
                liBranchGID = 17

                If liVendorGID = 0 Then _
                    lsErrNote &= " VENDOR NOT FOUND : "


                If lsErrNote <> "" Then GoTo NextFetch

                ' ECF HEADER GID FINDING 
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL = " SELECT ecf_gid FROM iem_trn_tecf "
                lsSQL &= " WHERE ecf_no='" & lsREFERENCENO & "' "

                liECFGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

                If liECFGID = 0 Then
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_trn_tecf VALUES (ecf_supplier_employee,ecf_supplier_gid,ecf_employee_gid,ecf_date,ecf_no,ecf_slno,ecf_create_mode,"
                    lsSQL &= " ecf_raiser,ecf_doctype_gid,ecf_docsubtype_gid,ecf_branch_gid,ecf_po_type,ecf_claim_month,ecf_currency_gid,ecf_currency_code,ecf_currency_rate,"
                    lsSQL &= " ecf_currency_amount,ecf_amount,ecf_reduced_amount,ecf_processed_amount,ecf_delmat_amount,ecf_person_count,ecf_despatch_date,ecf_courier_name,"
                    lsSQL &= " ecf_awb_no,ecf_remark,ecf_prev_status,ecf_status,ecf_all_status,ecf_queue_gid,ecf_queue_to_type,ecf_queue_to,ecf_action_by,ecf_action_date,"
                    lsSQL &= " ecf_amort_flag,ecf_amort_from,ecf_amort_to,ecf_amort_desc,ecf_amort_gid,ecf_urgent_flag,ecf_audit_update_flag,ecf_insert_by,ecf_insert_date,"
                    lsSQL &= " ecf_approved_date,ecf_auth_date,ecf_cancel_by,ecf_cancel_date,ecf_isremoved,ecf_travelpersoncount,ecf_description,ecf_payment_Nett,ecf_checklist,"
                    lsSQL &= " ecf_IsUpdated,ecf_paymentNett,ecf_printflag) "

                    lsSQL &= " ('S'," & liVendorGID & ", NULL,'2016-03-14 00:00:00','" & lsREFERENCENO & "',1,'S'," & liRaiserGID & ",3,4,17,'P','2016-03-11 00:00:00',99,'INR',1.00,"
                    lsSQL &= Val(lsAMOUNT) & ",0.00,0.00," & Val(lsBALANCEAMOUNT) & ",NULL, NULL,NULL,NULL,'" & lsNARRATION & "',NULL,65536,65536,0,'E',0," & liRaiserGID & ",'2016-03-14 15:33:00',"
                    lsSQL &= "'N',NULL,NULL,NULL,NULL,'N','N'," & liRaiserGID & ",'2016-03-14 15:32:00',NULL,NULL,NULL,NULL,'N',NULL,'" & lsDEPARTMENT & "','Y',NULL,0,NULL,'Y')"

                    lsErrMessage &= (loDBConnection.ExecuteNonQuerySQL(lsSQL))

                    ' ECF HEADER GID FINDING 
                    '-------------------------------------------------------------------------------------------------------------------------------
                    lsSQL = ""
                    lsSQL = " SELECT ecf_gid FROM iem_trn_tecf "
                    lsSQL &= " WHERE ecf_no='" & lsREFERENCENO & "' "

                    liECFGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)
                End If


                If liECFGID <> 0 Then
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_trn_tinvoice(invoice_ecf_gid, invoice_supplier_gid, invoice_employee_gid, invoice_type, invoice_date, invoice_service_month, "
                    lsSQL &= " invoice_no, invoice_desc, invoice_amount, invoice_wotax_amount, invoice_payment_nett, invoice_dedup_no, invoice_dedup_status, invoice_provision_flag, "
                    lsSQL &= " invoice_retention_flag, invoice_retention_rate, invoice_retention_amount, invoice_retention_exception, invoice_retention_curr_status, "
                    lsSQL &= " invoice_retention_status, invoice_retention_releaseon, invoice_isremoved, invoice_amort_flag, invoice_isupdated, Invoice_IsVerified, invoice_amort_gid, "
                    lsSQL &= " invoice_netpayable_amount) VALUES ("

                    lsSQL &= liECFGID & "," & liVendorGID & "," & liRaiserGID & ",'P','2016-03-14 15:32:00','2016-03-14 15:32:00','" & lsINVOICENO & "','"
                    lsSQL &= lsDEPARTMENT & "'," & Val(lsBALANCEAMOUNT) & ",0," & Val(lsBALANCEAMOUNT) & ",'" & lsINVOICENO & "',0,'Y','Y',100.00," & Val(lsBALANCEAMOUNT) & ","
                    lsSQL &= Val(lsBALANCEAMOUNT) & ",0,0,'2016-07-15 00:00:00','N',NULL,0,0,0,0) "

                    lsErrMessage &= (loDBConnection.ExecuteNonQuerySQL(lsSQL))

                    ' INVOICE GID FINDING 
                    '-------------------------------------------------------------------------------------------------------------------------------
                    lsSQL = ""
                    lsSQL = " SELECT MAX(invoice_gid) FROM iem_trn_tinvoice "
                    lsSQL &= " WHERE invoice_ecf_gid=" & liECFGID

                    liInvoiceGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)
                Else
                    lsErrNote &= " ECF# NOT INSERTED : "
                End If

                If liInvoiceGID <> 0 Then

                    ' DEBIT LINE GID FINDING 
                    '-------------------------------------------------------------------------------------------------------------------------------
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_trn_tecfdebitline (ecfdebitline_ecf_gid,ecfdebitline_invoice_gid,ecfdebitline_expnature_gid,ecfdebitline_expcat_gid,ecfdebitline_expsubcat_gid,ecfdebitline_gl_no,"
                    lsSQL &= " ecfdebitline_desc,ecfdebitline_period_from,ecfdebitline_period_to,ecfdebitline_fc_code,ecfdebitline_cc_code,ecfdebitline_product_code,ecfdebitline_ou_code,ecfdebitline_amount,"
                    lsSQL &= " ecfdebitline_isremoved,ecfdebitline_category_type,ecfdebitline_assetcategory_gid,ecfdebitline_assetsubcategory_gid,ecfdebitline_prodservice_gid,ecfdebitline_invoicepoitem_gid,"
                    lsSQL &= " ecfdebitline_ref_gid,ecfdebitline_ref_Rid) VALUES("

                    lsSQL &= liECFGID & "," & liInvoiceGID & ",65,363,25,441700002,'" & lsDEPARTMENT & "','2016-03-14 00:00:00',NULL,41,130,500,3503543," & lsAMOUNT & ",'N',NULL,0,0,0,0,0,0)"

                    lsErrMessage &= (loDBConnection.ExecuteNonQuerySQL(lsSQL))


                    lsSQL = ""
                    lsSQL = " SELECT MAX(ecfdebitline_gid) FROM iem_trn_tecfdebitline "
                    lsSQL &= " WHERE ecfdebitline_ecf_gid=" & liECFGID

                    liDebitLineGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


                    ' CREDIT LINE GID FINDING 
                    '-------------------------------------------------------------------------------------------------------------------------------
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_trn_tecfcreditline(ecfcreditline_ecf_gid,ecfcreditline_invoice_gid,ecfcreditline_pay_mode,ecfcreditline_ref_no,"
                    lsSQL &= " ecfcreditline_beneficiary,ecfcreditline_bank_gid,ecfcreditline_ifsc_code,ecfcreditline_gl_no,ecfcreditline_desc,ecfcreditline_amount,"
                    lsSQL &= " ecfcreditline_isremoved) VALUES("

                    lsSQL &= liECFGID & "," & liInvoiceGID & ",'RET','" & lsINVOICENO & "','" & lsVENDORNAME & "',0,NULL,'" & lsRETENTIONGL
                    lsSQL &= "','Invoice Retension Amount'," & lsAMOUNT & ",'N')"

                    lsErrMessage &= (loDBConnection.ExecuteNonQuerySQL(lsSQL))

                    lsSQL = ""
                    lsSQL = " SELECT MAX(creditline_gid) FROM iem_trn_tecfcreditline "
                    lsSQL &= " WHERE ecfcreditline_ecf_gid=" & liECFGID

                    liCreditLineGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

                    ' RETENTION LOG FINDING 
                    '-------------------------------------------------------------------------------------------------------------------------------
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_trn_tretentionlog(retention_date, retention_invoice_gid, retention_serialno, retention_amount,"
                    lsSQL &= " retention_rate,retention_releaseamount, retention_exception, retention_release_gid, retention_expiry, retention_status,"
                    lsSQL &= " retention_insertby, retention_inserton, retention_isactive, retention_remarks) VALUES('2016-03-14 16:54:00',"
                    lsSQL &= liInvoiceGID & ",1,100," & lsAMOUNT & ",0," & lsAMOUNT & ",0,'2016-07-14 00:00:00','Retention Book'," & liRaiserGID & ","
                    lsSQL &= "'2016-03-11 16:54:00','Y','ok')"

                    lsErrMessage &= (loDBConnection.ExecuteNonQuerySQL(lsSQL))

                    lsSQL = ""
                    lsSQL = " SELECT MAX(retention_gid) FROM iem_trn_tretentionlog "
                    lsSQL &= " WHERE retention_invoice_gid=" & liInvoiceGID

                    liRetentionGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)
                Else
                    lsErrNote &= " INVOICE# NOT INSERTED : "
                End If

NextFetch:
                If lsErrMessage = "" And lsErrNote = "" Then
                    liValid += 1
                Else
                    lsSQL = ""
                    lsSQL &= " DELETE FROM iem_trn_tretentionlog "
                    lsSQL &= " WHERE retentionlog_gid = " & liRetentionGID

                    lsErrMessage &= (loDBConnection.ExecuteNonQuerySQL(lsSQL))

                    lsSQL = ""
                    lsSQL &= " DELETE FROM iem_trn_tecfcreditline "
                    lsSQL &= " WHERE ecfcreditline_gid = " & liCreditLineGID

                    lsErrMessage &= (loDBConnection.ExecuteNonQuerySQL(lsSQL))

                    lsSQL = ""
                    lsSQL &= " DELETE FROM iem_trn_tecfdebitline "
                    lsSQL &= " WHERE ecfdebitline_gid = " & liDebitLineGID

                    lsErrMessage &= (loDBConnection.ExecuteNonQuerySQL(lsSQL))

                    lsSQL = ""
                    lsSQL &= " DELETE FROM iem_trn_tinvoice "
                    lsSQL &= " WHERE invoice_gid = " & liInvoiceGID

                    lsErrMessage &= (loDBConnection.ExecuteNonQuerySQL(lsSQL))

                    lsSQL = ""
                    lsSQL &= " DELETE FROM iem_trn_tecf "
                    lsSQL &= " WHERE ecf_gid = " & liECFGID

                    lsErrMessage &= (loDBConnection.ExecuteNonQuerySQL(lsSQL))


                    liError += 1

                    lobjErrorDatatable.Rows.Add()

                    With lobjErrorDatatable.Rows(lobjErrorDatatable.Rows.Count - 1)

                        .Item("SNO") = lsSNO
                        .Item("INWARD DATE") = lsINWARDDATE
                        .Item("REFERENCE NO") = lsREFERENCENO
                        .Item("AMOUNT") = lsAMOUNT
                        .Item("PAID AMOUNT") = lsPAIDAMOUNT
                        .Item("BALANCE AMOUNT") = lsBALANCEAMOUNT
                        .Item("RETENTION GL") = lsRETENTIONGL
                        .Item("VENDOR CODE") = lsVENDORCODE
                        .Item("VENDOR NAME") = lsVENDORNAME
                        .Item("INVOICE NO") = lsINVOICENO
                        .Item("BRANCH CODE") = lsBRANCHCODE
                        .Item("MODULE CODE") = lsMODULECODE
                        .Item("NARRATION") = lsNARRATION
                        .Item("DEPARTMENT") = lsDEPARTMENT
                        .Item("AGEING") = lsAGEING
                        .Item("Bucket") = lsBUCKET
                        .Item("Remark") = lsRemark

                        .Item("Error") = lsErrMessage
                        .Item("Error Note") = lsErrNote

                    End With

                End If

                txtStatus.Text = liValid & " Of " & lobjDataTable.Rows.Count & " Records Migrated...  Error - " & liError
                Application.DoEvents()
            Next

            If lobjErrorDatatable.Rows.Count > 0 Then
                PrintDGridviewXML(lobjErrorDatatable, Application.StartupPath & "\RET-MIG-ERR.xls")
                MessageBox.Show("Descrepance Records Spooled @ " & Application.StartupPath & "\RET-MIG-ERR.xls")
            Else
                MessageBox.Show("Imported successfully...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If


        Catch ex As Exception
            grpMain.Enabled = True
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub




    Public Sub WO_Data_Migration()
        Try
            Dim lsSQL As String = ""
            Dim lsFileName As String = ""
            Dim lsResult As String = ""
            Dim liDumpGid As Integer = 0
            Dim lobjOleDbConnection As New OleDb.OleDbConnection
            Dim lobjDataAdapter As OleDb.OleDbDataAdapter
            Dim lobjDataTable As DataTable
            Dim lobjDataSet As New DataSet
            Dim liError As Integer
            Dim liValid As Integer
            Dim liBranchGID As Integer

            Dim lsErrMessage As String = ""
            Dim lsErrNote As String = ""



            lsFileName = Mid(txtFileName.Text.Trim, InStrRev(txtFileName.Text.Trim, "\") + 1)

            Try
                If lobjOleDbConnection.State = 1 Then lobjOleDbConnection.Close()
                With lobjOleDbConnection

                    If Microsoft.VisualBasic.Right(lsFileName, 4) = "xlsx" Then
                        'read a 2007 file   
                        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                                             txtFileName.Text.Trim & ";" + ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
                    Else
                        'read a 97-2003 file   
                        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" & _
                                                txtFileName.Text.Trim & ";" + "Extended Properties=Excel 8.0;"
                    End If

                    .Open()

                End With

            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                lobjOleDbConnection.Close()
                Exit Sub
            End Try

            lsSQL = "SELECT * FROM [" & cmbSheet.Text & "$]"

            lobjDataAdapter = New OleDb.OleDbDataAdapter(lsSQL, lobjOleDbConnection)
            lobjDataAdapter.Fill(lobjDataSet, "DATA")
            lobjDataTable = lobjDataSet.Tables("DATA")
            lobjOleDbConnection.Close()

            Dim liIndex As Integer

            For liIndex = 0 To lobjDataTable.Columns.Count - 1
                If Not InStr(fsColumnHeaders.ToLower.Trim, "," & lobjDataTable.Columns(liIndex).ColumnName.ToString.Trim.ToLower & ",") > 0 Then
                    MessageBox.Show("WO Dump is not in the correct format...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If
            Next

            Dim lsWONumber As String
            Dim lsCBF_header_CBFNo As String
            Dim lsCBFDetail_Sno As String
            Dim lsWODate As String
            Dim lsWORaisor As String
            Dim lsRequestFor As String
            Dim lsITType As String
            Dim lsVendorName As String
            Dim lsVendor_note As String
            Dim lsHeaderTotal As String
            Dim lsFrequencyType As String
            Dim lsFrom_month As String
            Dim lsTo_month As String
            Dim lsType As String
            Dim lsTerm_and_condition As String
            Dim lsAdditional_term_and_condition As String
            Dim lsProduct_service As String
            Dim lsDescription As String
            Dim lsService_month As String
            Dim lsPercentage As String
            Dim lsDetails_total As String
            Dim lsQty As String
            Dim lsLegacyBranch As String

            Dim liSerialNo As String

            Dim liCBFGID As Integer
            Dim liProductGID As Integer
            Dim liUOMGID As Integer
            Dim liVendorGID As Integer
            Dim liProdServiceGID As Integer
            Dim liCBFDetailsGID As Integer
            Dim liRaisorGID As Integer
            Dim liPOGID As Integer
            Dim liPODETGID As Integer
            Dim liPOShipmentGID As Integer
            Dim liGRNReleaseforPOGID As Integer
            Dim liGRNInwardGID As Integer


            Dim lobjErrorDatatable As New DataTable

            With lobjErrorDatatable
                .Columns.Add("PO Number")
                .Columns.Add("CBF_header_CBFNo")
                .Columns.Add("CBFDetail_Sno")
                .Columns.Add("Date")
                .Columns.Add("Raisor")
                .Columns.Add("RequestFor")
                .Columns.Add("IT Type")
                .Columns.Add("Vendor Name")
                .Columns.Add("Vendor_note")
                .Columns.Add("Header Total")
                .Columns.Add("Frequency Type")
                .Columns.Add("From_month")
                .Columns.Add("To_month")
                .Columns.Add("Type")
                .Columns.Add("Term_and_condition")

                .Columns.Add("Additional_term_and_condition")

                .Columns.Add("Product_service")
                .Columns.Add("Description")
                .Columns.Add("Service_month")
                .Columns.Add("Percentage")
                .Columns.Add("Details_total")
                .Columns.Add("PO Detail Qty")
                .Columns.Add("Legacy Branch")

                .Columns.Add("Error")
                .Columns.Add("Error Note")
            End With

            For i As Integer = 0 To lobjDataTable.Rows.Count - 1
                lsErrMessage = ""
                lsErrNote = ""

                With lobjDataTable.Rows(i)

                    lsWONumber = QuoteFilter(.Item("PO Number").ToString)
                    lsCBF_header_CBFNo = QuoteFilter(.Item("CBF_header_CBFNo").ToString)
                    lsCBFDetail_Sno = QuoteFilter(.Item("CBFDetail_Sno").ToString)
                    lsWODate = QuoteFilter(.Item("Date").ToString)
                    lsWORaisor = QuoteFilter(.Item("Raisor").ToString)
                    lsRequestFor = QuoteFilter(.Item("RequestFor").ToString)
                    lsITType = QuoteFilter(.Item("IT Type").ToString)
                    lsVendorName = QuoteFilter(.Item("Vendor Name").ToString)
                    lsVendor_note = QuoteFilter(.Item("Vendor_note").ToString)
                    lsHeaderTotal = QuoteFilter(.Item("Header Total").ToString)
                    lsFrequencyType = QuoteFilter(.Item("Frequency Type").ToString)
                    lsFrom_month = QuoteFilter(.Item("From_month").ToString)
                    lsTo_month = QuoteFilter(.Item("To_month").ToString)
                    lsType = QuoteFilter(.Item("Type").ToString)
                    lsTerm_and_condition = QuoteFilter(.Item("Term_and_condition").ToString)

                    lsAdditional_term_and_condition = QuoteFilter(.Item("Additional_term_and_condition").ToString)

                    lsProduct_service = QuoteFilter(.Item("Product_service").ToString)
                    lsDescription = QuoteFilter(.Item("Description").ToString)
                    lsService_month = QuoteFilter(.Item("Service_month").ToString)
                    lsPercentage = QuoteFilter(.Item("Percentage").ToString)
                    lsDetails_total = QuoteFilter(.Item("Details_total").ToString)

                    lsQty = QuoteFilter(.Item("PO Detail Qty").ToString)
                    lsLegacyBranch = QuoteFilter(.Item("Legacy Branch").ToString)

                End With

                If lsWODate <> "" Then
                    If IsDate(lsWODate) Then
                        lsWODate = "'" & Format(CDate(lsWODate), "yyyy-MM-dd") & "'"
                    Else
                        lsWODate = "NULL"
                    End If
                Else
                    lsWODate = "NULL"
                End If

                liCBFGID = 0
                liProductGID = 0
                liUOMGID = 0
                liVendorGID = 0
                liProdServiceGID = 0
                liCBFDetailsGID = 0
                liRaisorGID = 0
                liPOGID = 0
                liPODETGID = 0
                liPOShipmentGID = 0
                liGRNReleaseforPOGID = 0
                liGRNInwardGID = 0
                liSerialNo = 0

                ' UOM GID 
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL = " SELECT uom_gid FROM iem_mst_tuom "
                lsSQL &= " WHERE uom_code='NO' "

                liUOMGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


                ' PRODUCT / SERVICE GID 
                '-------------------------------------------------------------------------------------------------------------------------------
                If InStr(lsProduct_service, "-") = 0 Then _
                    lsProduct_service &= " - "

                lsSQL = ""
                lsSQL = " SELECT prodservice_gid FROM fb_mst_tprodservice "
                lsSQL &= " WHERE prodservice_code='" & Mid(lsProduct_service, 1, InStr(lsProduct_service, "-") - 1).Trim & "' "
                lsSQL &= " AND prodservice_isremoved='N' "

                liProductGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

                lsSQL = ""
                lsSQL = " SELECT prodservice_prodservicegid FROM fb_mst_tprodservice "
                lsSQL &= " WHERE prodservice_code='" & Mid(lsProduct_service, 1, InStr(lsProduct_service, "-") - 1).Trim & "' "
                lsSQL &= " AND prodservice_isremoved='N' "

                liProdServiceGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


                ' SUPPLIER GID 
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL = " SELECT supplierheader_gid FROM asms_trn_tsupplierheader "
                lsSQL &= " WHERE supplierheader_name='" & lsVendorName & "' "  'Mid(lsVendorName, 1, InStr(lsVendorName, "-") - 1).Trim

                liVendorGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


                ' EMPLOYEE GID 
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL = " SELECT employee_gid FROM iem_mst_temployee "
                lsSQL &= " WHERE employee_code='" & lsWORaisor & "' "                           'Mid(lsWORaisor, 1, InStr(lsWORaisor, "-") - 1).Trim

                liRaisorGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

                ' BRANCH GID 
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL = " SELECT branch_gid FROM iem_mst_tbranch "
                lsSQL &= " WHERE branch_legacy_code='" & Microsoft.VisualBasic.Right(("0000" & lsLegacyBranch), 4) & "'"

                liBranchGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

                If liBranchGID = 0 Then
                    lsSQL = ""
                    lsSQL = " SELECT branch_gid FROM iem_mst_tbranch "
                    lsSQL &= " WHERE LEFT(branch_name,4)='" & Microsoft.VisualBasic.Right(("0000" & lsLegacyBranch), 4) & "'"

                    liBranchGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)
                End If


                If liProductGID = 0 Then _
                    lsErrNote &= " PRODUCT NOT FOUND : "

                If liProdServiceGID = 0 Then _
                    lsErrNote &= " PRODUCT SERVICE NOT FOUND : "

                If liVendorGID = 0 Then _
                    lsErrNote &= " VENDOR NOT FOUND : "

                If liRaisorGID = 0 Then _
                    lsErrNote &= " RAISER NOT FOUND : "

                If liBranchGID = 0 Then _
                    lsErrNote &= " BRANCH NOT FOUND : "

                If lsErrNote <> "" Then GoTo NextFetch

                ' CBF HEADER GID FINDING 
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL = " SELECT cbfheader_gid FROM fb_trn_tcbfheader "
                lsSQL &= " WHERE cbfheader_cbfno='" & lsCBF_header_CBFNo & "' "

                liCBFGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

                If liCBFGID = 0 Then

                    ' INSERTING DATA INTO CBF HEADER 
                    '-------------------------------------------------------------------------------------------------------------------------------
                    lsSQL = ""
                    lsSQL &= " INSERT INTO fb_trn_tcbfheader(cbfheader_cbfno,cbfheader_cbfobf_flag,cbfheader_date,cbfheader_enddate,cbfheader_projectowner,cbfheader_branch_gid,"
                    lsSQL &= " cbfheader_mode,cbfheader_prpar_gid,cbfheader_approvaltype,cbfheader_isbudgeted,cbfheader_Devi_amt,cbfheader_cbfamt,cbfheader_desc,"
                    lsSQL &= " cbfheader_rasier_gid,cbfheader_requestfor_gid,cbfheader_requesttype,cbfheader_remarks,cbfheader_budgetowner_gid,cbfheader_status)  "

                    lsSQL &= " VALUES('" & lsCBF_header_CBFNo & "','O'," & lsWODate & ",'2017-12-31',0," & liBranchGID & ",'PAR',0,'R','Y',0,0,'WO-OBF UAT DATAMIGRATION',0,0,'','',0,5) "

                    lsErrMessage &= (loDBConnection.ExecuteNonQuerySQL(lsSQL))

                    lsSQL = ""
                    lsSQL = " SELECT cbfheader_gid FROM fb_trn_tcbfheader "
                    lsSQL &= " WHERE cbfheader_cbfno='" & lsCBF_header_CBFNo & "' "

                    liCBFGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

                End If

                If liCBFGID <> 0 Then
                    ' INSERTING DATA INTO CBF DETAILS 
                    '-------------------------------------------------------------------------------------------------------------------------------
                    lsSQL = ""
                    lsSQL &= " INSERT INTO fb_trn_tcbfdetails(cbfheader_cbfobf_flag,cbfdetails_cbfhead_gid,cbfdetails_parprdesc,cbfdetails_year,cbfdetails_prod_gid,"
                    lsSQL &= " cbfdetails_desc,cbfdetails_uom_gid,cbfdetails_qty,cbfdetails_unitprice,cbfdetails_totalamt,cbfdetails_remarks,cbfdetails_chartofacc,cbfdetails_fccc,"
                    lsSQL &= " cbfdetails_budgetline,cbfdetails_budgetowner_gid,cbfdetails_vendor_gid,cbfdetails_prpardel_gid,cbfdetails_prodservgrp_gid,"
                    lsSQL &= " cbfdetails_sno) "

                    lsSQL &= " VALUES ('O'," & liCBFGID & ",'MIG',''," & liProductGID & ",'" & lsDescription & "'," & liUOMGID & "," & Val(lsQty)
                    lsSQL &= "," & Val(lsDetails_total) & "," & Val(lsDetails_total) & ",'" & lsDescription & "','1111','11103',0,0," & liVendorGID & ",0," & liProdServiceGID
                    lsSQL &= "," & Val(lsCBFDetail_Sno) & ")"

                    lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)


                    lsSQL = ""
                    lsSQL = " SELECT cbfdetails_gid FROM fb_trn_tcbfdetails "
                    lsSQL &= " WHERE cbfdetails_cbfhead_gid='" & liCBFGID & "' AND cbfdetails_sno = " & Val(lsCBFDetail_Sno)

                    liCBFDetailsGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)
                Else
                    lsErrNote &= " CBF# NOT FOUND : "
                End If

                If liCBFDetailsGID <> 0 Then

                    lsSQL = ""
                    lsSQL = " SELECT poheader_gid FROM fb_trn_tpoheader "
                    lsSQL &= " WHERE poheader_pono='" & lsWONumber & "' "

                    liPOGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

                    If liPOGID = 0 Then
                        ' INSERTING DATA INTO PO HEADER
                        '-------------------------------------------------------------------------------------------------------------------------------
                        lsSQL = ""
                        lsSQL &= " INSERT INTO fb_trn_tpoheader (poheader_pono, poheader_date, poheader_raisor_gid, poheader_ittype, poheader_requestfor,  poheader_vendor_gid, "
                        lsSQL &= " poheader_vendor_note, poheader_over_total, poheader_frequency_gid, poheader_from_month, poheader_to_month, poheader_type, poheader_status, "
                        lsSQL &= " poheader_termcond_gid, poheader_add_termandcond, poheader_currentapprovalstage )  "

                        lsSQL &= " VALUES ('" & lsWONumber & "'," & lsWODate & "," & liRaisorGID & ",'" & Mid(lsITType, 1, 1) & "',"
                        lsSQL &= " '" & IIf(Mid(lsRequestFor, 1, 1) = "P", 3, 4) & "'," & liVendorGID & ",'" & lsVendor_note & "'," & Val(lsDetails_total) & ",1,"
                        lsSQL &= "'" & lsFrom_month & "','" & lsTo_month & "','" & lsType & "',5,0,'" & lsTerm_and_condition & "',Null)"

                        lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)

                        lsSQL = ""
                        lsSQL = " SELECT poheader_gid FROM fb_trn_tpoheader "
                        lsSQL &= " WHERE poheader_pono='" & lsWONumber & "' "

                        liPOGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)
                    End If
                Else
                    If liCBFGID <> 0 Then lsErrNote &= " CBF ITEM NOT FOUND : "
                End If


                If liPOGID <> 0 Then

                    lsSQL = ""
                    lsSQL = " SELECT COUNT(*) as CNT FROM fb_trn_tpodetails WHERE podetails_pohead_gid = " & liPOGID

                    liSerialNo = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString) + 1

                    ' INSERTING DATA INTO PO DETAILS
                    '-------------------------------------------------------------------------------------------------------------------------------
                    lsSQL = ""

                    lsSQL &= " INSERT INTO fb_trn_tpodetails ( podetails_pohead_gid, podetails_prodservice_gid, podetails_desc, podetails_serv_month, podetails_unitprice, "
                    lsSQL &= " podetails_total, podetails_cbfdet_gid, podetails_qty, podetails_uom_gid, podetails_base_amt,podetails_sno) VALUES(" & liPOGID & "," & liProductGID & ",'" & lsDescription & "','" & lsService_month & "',"
                    lsSQL &= lsDetails_total & "," & lsDetails_total & "," & liCBFDetailsGID & "," & lsQty & "," & liUOMGID & "," & lsDetails_total & "," & liSerialNo & ")"

                    lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)


                    lsSQL = ""
                    lsSQL = " SELECT podetails_gid FROM fb_trn_tpodetails "
                    lsSQL &= " WHERE podetails_pohead_gid='" & liPOGID & "' AND podetails_sno = " & liSerialNo
                    'lsSQL &= " AND podetails_prodservice_gid = " & liProductGID
                    'lsSQL &= " AND podetails_desc='" & lsDescription & "' "
                    'lsSQL &= " AND podetails_qty= " & lsQty

                    liPODETGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)
                Else
                    lsErrNote &= " PO# NOT FOUND : "
                End If

                If liPODETGID <> 0 Then

                    ' INSERTING DATA INTO PO SHIPMENT DETAILS
                    '-------------------------------------------------------------------------------------------------------------------------------
                    lsSQL = ""
                    lsSQL &= " INSERT INTO fb_trn_tposhipment ( "
                    lsSQL &= " poshipment_podet_gid,poshipment_type_gid,poshipment_branch_gid,"
                    lsSQL &= " poshipment_qty,poshipment_isremoved,poshipment_empgid,"
                    lsSQL &= " poshipment_isamended,poshipment_remarks) VALUES("
                    lsSQL &= liPODETGID & ",1," & liBranchGID & "," & lsQty & ",'N'," & liRaisorGID & ",'N','Test Wo Migration') "

                    lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)

                    lsSQL = ""
                    lsSQL = " SELECT poshipment_gid FROM fb_trn_tposhipment "
                    lsSQL &= " WHERE poshipment_podet_gid=" & liPODETGID

                    liPOShipmentGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


                    ' INSERTING DATA INTO GRN RELEASE FOR PO
                    '-------------------------------------------------------------------------------------------------------------------------------
                    lsSQL = ""
                    lsSQL &= "INSERT INTO fb_trn_tgrnreleaseforpo ("
                    lsSQL &= " grnreleaseforpo_podet_gid,grnreleaseforpo_released_qty,grnreleaseforpo_isremoved,"
                    lsSQL &= " grnreleaseforpo_balanceqty,grnreleaseforpo_poshipment_gid,grnreleaseforpo_branch_type,"
                    lsSQL &= " grnreleaseforpo_released_date,grnreleaseforpo_releasedby,inward_release) VALUES("
                    lsSQL &= liPODETGID & "," & lsQty & ",'N',0," & liPOShipmentGID & ",'D','" & Format(Now, "yyyy-MM-dd") & "', "
                    lsSQL &= liRaisorGID & ",Null)"

                    lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)

                    lsSQL = ""
                    lsSQL = " SELECT grnreleaseforpo_gid FROM fb_trn_tgrnreleaseforpo "
                    lsSQL &= " WHERE grnreleaseforpo_podet_gid=" & liPODETGID
                    lsSQL &= " AND grnreleaseforpo_poshipment_gid=" & liPOShipmentGID

                    liGRNReleaseforPOGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


                    ' INSERTING DATA INTO GRN INWARD
                    '-------------------------------------------------------------------------------------------------------------------------------
                    lsSQL = ""
                    lsSQL &= " INSERT INTO fb_trn_tgrninwrdheader(grninwrdheader_refno,grninwrdheader_grndatetime,grninwardheader_poheader,grninwrdheader_rasiergid,"
                    lsSQL &= " grninwrdheader_dcno,grninwrdheader_invoiceno,grninwrdheader_remarks,grninwrdheader_isremoved,grninwrdheader_status,grninwrdheader_branch_type,"
                    lsSQL &= " grniwrdheader_capstatus) VALUES("
                    lsSQL &= "'GRN" & Format(Now, "yyMMdd") & Format(i, "0000") & "',SYSDATETIME()," & liPOGID & "," & liRaisorGID & ","
                    lsSQL &= "'12345','12345','Test Wo Migration','N','5','D','') "

                    lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)

                    lsSQL = ""
                    lsSQL = " SELECT grninwrdheader_gid FROM fb_trn_tgrninwrdheader "
                    lsSQL &= " WHERE grninwardheader_poheader=" & liPOGID

                    liGRNInwardGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


                    ' INSERTING DATA INTO GRN INWARD DETAILS 
                    '-------------------------------------------------------------------------------------------------------------------------------
                    lsSQL = ""
                    lsSQL &= " INSERT INTO fb_trn_tgrninwrddet( "
                    lsSQL &= " grninwrddet_grnreleforpo_gid,grninwrddet_reced_qty,grninwrddet_reced_date,"
                    lsSQL &= " grninwrddet_isremoved,grninwrddet_grninwrdhead_gid,grninwrddet_assetsrlno, "
                    lsSQL &= " grninwrddet_puttousedatetime,grninwrddet_mft_name,grninwrddet_capstatus ) VALUES( "
                    lsSQL &= liGRNReleaseforPOGID & "," & lsQty & ",SYSDATETIME(),'N'," & liGRNInwardGID & ",'',SYSDATETIME(),'',Null) "

                    lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)


                    ' INSERTING DATA INTO GRN CONFIRMATION 
                    '-------------------------------------------------------------------------------------------------------------------------------
                    lsSQL = ""
                    lsSQL &= " INSERT INTO fb_trn_tgrnconfrm(grnconfrm_grninwrdheader_gid,grnconfrm_remarks,grnconfrm_date,grnconfrm_confirmby,grnconfrm_status,"
                    lsSQL &= " grnconfrm_isremoved)values(" & liGRNInwardGID & ",'Test Migration',SYSDATETIME() ," & liRaisorGID & ",'5','N')"

                    lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)
                Else
                    lsErrNote &= " PO ITEM NOT FOUND : "
                End If

NextFetch:
                If lsErrMessage = "" And lsErrNote = "" Then
                    liValid += 1
                Else
                    liError += 1

                    lobjErrorDatatable.Rows.Add()
                    With lobjErrorDatatable.Rows(lobjErrorDatatable.Rows.Count - 1)
                        .Item("PO Number") = lsWONumber
                        .Item("CBF_header_CBFNo") = lsCBF_header_CBFNo
                        .Item("CBFDetail_Sno") = lsCBFDetail_Sno
                        .Item("Date") = lsWODate
                        .Item("Raisor") = lsWORaisor
                        .Item("RequestFor") = lsRequestFor
                        .Item("IT Type") = lsITType
                        .Item("Vendor Name") = lsVendorName
                        .Item("Vendor_note") = lsVendor_note
                        .Item("Header Total") = lsHeaderTotal
                        .Item("Frequency Type") = lsFrequencyType
                        .Item("From_month") = lsFrom_month
                        .Item("To_month") = lsTo_month
                        .Item("Type") = lsType
                        .Item("Term_and_condition") = lsTerm_and_condition

                        .Item("Additional_term_and_condition") = lsDetails_total

                        .Item("Product_service") = lsProduct_service
                        .Item("Description") = lsDescription
                        .Item("Service_month") = lsService_month
                        .Item("Percentage") = lsPercentage
                        .Item("Details_total") = lsDetails_total

                        .Item("PO Detail Qty") = lsQty
                        .Item("Error") = lsErrMessage
                        .Item("Error Note") = lsErrNote
                    End With

                End If

                txtStatus.Text = liValid & " Of " & lobjDataTable.Rows.Count & " Records Migrated...  Error - " & liError
                Application.DoEvents()
            Next

            If lobjErrorDatatable.Rows.Count > 0 Then
                PrintDGridviewXML(lobjErrorDatatable, Application.StartupPath & "\WO-MIG-ERR.xls")
                MessageBox.Show("Descrepance Records Spooled @ " & Application.StartupPath & "\WO-MIG-ERR.xls")
            Else
                MessageBox.Show("Imported successfully...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If


        Catch ex As Exception
            grpMain.Enabled = True
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub PO_Data_Migration()
        'Try
        Dim lsSQL As String = ""
        Dim lsFileName As String = ""
        Dim lsResult As String = ""
        Dim liDumpGid As Integer = 0
        Dim lobjOleDbConnection As New OleDb.OleDbConnection
        Dim lobjDataAdapter As OleDb.OleDbDataAdapter
        Dim lobjDataTable As DataTable
        Dim lobjDataSet As New DataSet
        Dim liError As Integer
        Dim liValid As Integer

        Dim lsErrMessage As String = ""
        Dim lsErrNote As String = ""


        lsFileName = Mid(txtFileName.Text.Trim, InStrRev(txtFileName.Text.Trim, "\") + 1)

        Try
            If lobjOleDbConnection.State = 1 Then lobjOleDbConnection.Close()
            With lobjOleDbConnection

                If Microsoft.VisualBasic.Right(lsFileName, 4) = "xlsx" Then
                    'read a 2007 file   
                    .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                                         txtFileName.Text.Trim & ";" + ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
                Else
                    'read a 97-2003 file   
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" & _
                                            txtFileName.Text.Trim & ";" + "Extended Properties=Excel 8.0;"
                End If

                .Open()

            End With

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            lobjOleDbConnection.Close()
            Exit Sub
        End Try

        lsSQL = "SELECT * FROM [" & cmbSheet.Text & "$]"

        lobjDataAdapter = New OleDb.OleDbDataAdapter(lsSQL, lobjOleDbConnection)
        lobjDataAdapter.Fill(lobjDataSet, "DATA")
        lobjDataTable = lobjDataSet.Tables("DATA")
        lobjOleDbConnection.Close()

        Dim liIndex As Integer

        For liIndex = 0 To lobjDataTable.Columns.Count - 1
            If Not InStr(fsColumnHeaders.ToLower.Trim, "," & lobjDataTable.Columns(liIndex).ColumnName.ToString.Trim.ToLower & ",") > 0 Then
                MessageBox.Show("PO Dump is not in the correct format...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Next

        'Sl.NO	poheader_pono	podetails_cbfheader_cbfno	CBFDetail_Sno	poheader_date	poheader_enddate	poheader_raisor	poheader_projectmanager	
        'poheader_requestfor	poheader_ittype	poheader_vendor	poheader_vendor_note	poheader_over_total	poheader_type	poheader_termcond	poheader_add_termandcond	
        'podetails_prodservice	podetails_desc	podetails_uom	podetails_qty	podetails_unitprice	podetails_discount	podetails_base_amt	podetails_tax1	
        'podetails_tax2  podetails_tax3 podetails_total 

        Dim lsSLNO As String
        Dim lsPONumber As String
        Dim lsPOCBFNumber As String
        Dim lsPOCBFSLNumber As String

        Dim lsPODate As String
        Dim lsPOEndDate As String
        Dim lsPORaisor As String
        Dim lsPOProjectManager As String

        Dim lsPORequestFor As String
        Dim lsPOITType As String
        Dim lsPOVendor As String
        Dim lsPOVendorNote As String

        Dim lsPOOverTotal As String
        Dim lsPOType As String
        Dim lsPOTermCond As String
        Dim lsPOAdditionalTerm As String

        Dim lsProduct_service As String
        Dim lsDescription As String
        Dim lsUOM As String
        Dim lsQty As String

        Dim lsUnitPrice As String
        Dim lsDiscount As String
        Dim lsBaseAmount As String

        Dim lsTax1 As String
        Dim lsTax2 As String
        Dim lsTax3 As String
        Dim lsTotal As String

        Dim lsStatus As String


        Dim liCBFGID As Integer
        Dim liCBFDetailsGID As Integer
        Dim liProductGID As Integer
        Dim liUOMGID As Integer
        Dim liVendorGID As Integer
        Dim liProdServiceGID As Integer
        Dim liRaisorGID As Integer
        Dim liPOGID As Integer
        Dim liPODETGID As Integer

        Dim liProjectOwnerGID As Integer
        Dim liBudgetOwnerGID As Integer

        Dim lobjErrorDatatable As New DataTable

        With lobjErrorDatatable
            .Columns.Add("SlNO")
            .Columns.Add("poheader_pono")
            .Columns.Add("podetails_cbfheader_cbfno")
            .Columns.Add("CBFDetail_Sno")

            .Columns.Add("poheader_date")
            .Columns.Add("poheader_enddate")
            .Columns.Add("poheader_raisor")
            .Columns.Add("poheader_projectmanager")

            .Columns.Add("poheader_requestfor")
            .Columns.Add("poheader_ittype")
            .Columns.Add("poheader_vendor")
            .Columns.Add("poheader_vendor_note")

            .Columns.Add("poheader_over_total")
            .Columns.Add("poheader_type")
            .Columns.Add("poheader_termcond")
            .Columns.Add("poheader_add_termandcond")

            .Columns.Add("podetails_prodservice")
            .Columns.Add("podetails_desc")
            .Columns.Add("podetails_uom")
            .Columns.Add("podetails_qty")

            .Columns.Add("podetails_unitprice")
            .Columns.Add("podetails_discount")
            .Columns.Add("podetails_base_amt")
            .Columns.Add("podetails_tax1")

            .Columns.Add("podetails_tax2")
            .Columns.Add("podetails_tax3")
            .Columns.Add("podetails_total")

            .Columns.Add("Status")

            .Columns.Add("Error")
            .Columns.Add("Error Note")
        End With

        For i As Integer = 0 To lobjDataTable.Rows.Count - 1
            lsErrMessage = ""
            lsErrNote = ""

            With lobjDataTable.Rows(i)
                lsSLNO = QuoteFilter(.Item("SlNO").ToString)
                lsPONumber = QuoteFilter(.Item("poheader_pono").ToString)
                lsPOCBFNumber = QuoteFilter(.Item("podetails_cbfheader_cbfno").ToString)
                lsPOCBFSLNumber = QuoteFilter(.Item("CBFDetail_Sno").ToString)

                lsPODate = QuoteFilter(.Item("poheader_date").ToString)
                lsPOEndDate = QuoteFilter(.Item("poheader_enddate").ToString)
                lsPORaisor = QuoteFilter(.Item("poheader_raisor").ToString)
                lsPOProjectManager = QuoteFilter(.Item("poheader_projectmanager").ToString)

                lsPORequestFor = QuoteFilter(.Item("poheader_requestfor").ToString)
                lsPOITType = QuoteFilter(.Item("poheader_ittype").ToString)
                lsPOVendor = QuoteFilter(.Item("poheader_vendor").ToString)
                lsPOVendorNote = QuoteFilter(.Item("poheader_vendor_note").ToString)

                lsPOOverTotal = QuoteFilter(.Item("poheader_over_total").ToString)
                lsPOType = QuoteFilter(.Item("poheader_type").ToString)
                lsPOTermCond = QuoteFilter(.Item("poheader_termcond").ToString)
                lsPOAdditionalTerm = QuoteFilter(.Item("poheader_add_termandcond").ToString)

                lsProduct_service = QuoteFilter(.Item("podetails_prodservice").ToString)
                lsDescription = QuoteFilter(.Item("podetails_desc").ToString)
                lsUOM = QuoteFilter(.Item("podetails_uom").ToString)
                lsQty = QuoteFilter(.Item("podetails_qty").ToString)

                lsUnitPrice = QuoteFilter(.Item("podetails_unitprice").ToString)
                lsDiscount = QuoteFilter(.Item("podetails_discount").ToString)
                lsBaseAmount = QuoteFilter(.Item("podetails_base_amt").ToString)
                lsTax1 = QuoteFilter(.Item("podetails_tax1").ToString)

                lsTax2 = QuoteFilter(.Item("podetails_tax2").ToString)
                lsTax3 = QuoteFilter(.Item("podetails_tax3").ToString)
                lsTotal = QuoteFilter(.Item("podetails_total").ToString)
                lsStatus = QuoteFilter(.Item("Status").ToString)

            End With

            If lsPODate <> "" Then
                If IsDate(lsPODate) Then
                    lsPODate = "'" & Format(CDate(lsPODate), "yyyy-MM-dd") & "'"
                Else
                    lsPODate = "NULL"
                End If
            Else
                lsPODate = "NULL"
            End If

            If lsPOEndDate <> "" Then
                If IsDate(lsPOEndDate) Then
                    lsPOEndDate = "'" & Format(CDate(lsPOEndDate), "yyyy-MM-dd") & "'"
                Else
                    lsPOEndDate = "NULL"
                End If
            Else
                lsPOEndDate = "NULL"
            End If

            liCBFGID = 0
            liProductGID = 0
            liUOMGID = 0
            liVendorGID = 0
            liProdServiceGID = 0
            liCBFDetailsGID = 0
            liRaisorGID = 0
            liPOGID = 0
            liPODETGID = 0

            ' UOM GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT uom_gid FROM iem_mst_tuom "
            lsSQL &= " WHERE LOWER(uom_code)='" & lsUOM.ToLower.Trim & "'"

            liUOMGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            ' PRODUCT / SERVICE GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            If InStr(lsProduct_service, "-") = 0 Then _
                lsProduct_service &= " - "

            lsSQL = ""
            lsSQL = " SELECT prodservice_gid FROM fb_mst_tprodservice "
            lsSQL &= " WHERE prodservice_code='" & Mid(lsProduct_service, 1, InStr(lsProduct_service, "-") - 1).Trim & "' "

            liProductGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            lsSQL = ""
            lsSQL = " SELECT prodservice_prodservicegid FROM fb_mst_tprodservice "
            lsSQL &= " WHERE prodservice_code='" & Mid(lsProduct_service, 1, InStr(lsProduct_service, "-") - 1).Trim & "' "

            liProdServiceGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            ' SUPPLIER GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT supplierheader_gid FROM asms_trn_tsupplierheader "
            lsSQL &= " WHERE supplierheader_name='" & lsPOVendor & "' "

            liVendorGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            ' EMPLOYEE GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT employee_gid FROM iem_mst_temployee "
            lsSQL &= " WHERE employee_code='" & lsPORaisor.Trim & "' "

            liRaisorGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            ' PROJECT OWNER GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT employee_gid FROM iem_mst_temployee "
            lsSQL &= " WHERE employee_code='" & Mid(lsPOProjectManager, 1, InStr(lsPOProjectManager, "-") - 1).Trim & "' "

            liProjectOwnerGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            lsSQL = ""
            lsSQL = " SELECT projectowner_gid FROM iem_mst_tprojectowner "
            lsSQL &= " WHERE projectowner_employeegid=" & liProjectOwnerGID & " "

            liProjectOwnerGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            ' CBF HEADER GID FINDING 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT cbfheader_gid FROM fb_trn_tcbfheader "
            lsSQL &= " WHERE cbfheader_cbfno='" & lsPOCBFNumber & "' "

            liCBFGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            lsSQL = ""
            lsSQL = " SELECT cbfdetails_gid FROM fb_trn_tcbfdetails "
            lsSQL &= " WHERE cbfdetails_cbfhead_gid='" & liCBFGID & "' AND cbfdetails_sno = " & Val(lsPOCBFSLNumber)

            liCBFDetailsGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            If liProductGID = 0 Then _
                lsErrNote &= " PRODUCT NOT FOUND : "

            If liProdServiceGID = 0 Then _
                lsErrNote &= " PRODUCT SERVICE NOT FOUND : "

            If liVendorGID = 0 Then _
                lsErrNote &= " VENDOR NOT FOUND : "

            If liRaisorGID = 0 Then _
                lsErrNote &= " RAISER NOT FOUND : "

            If liProjectOwnerGID = 0 Then _
                lsErrNote &= " PROJECT OWNER NOT FOUND : "

            If liCBFGID = 0 Then _
                lsErrNote &= " CBF# NOT FOUND : "

            If liCBFDetailsGID = 0 Then _
                lsErrNote &= " CBF ITEM NOT FOUND : "


            If lsErrNote <> "" Then GoTo NextFetch

            lsSQL = ""
            lsSQL = " SELECT poheader_gid FROM fb_trn_tpoheader "
            lsSQL &= " WHERE poheader_pono='" & lsPONumber & "' "

            liPOGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            If liPOGID = 0 Then
                ' INSERTING DATA INTO PO HEADER
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL &= " INSERT INTO fb_trn_tpoheader (poheader_pono, poheader_date, poheader_enddate, poheader_raisor_gid, poheader_ittype, poheader_requestfor,  poheader_vendor_gid, "
                lsSQL &= " poheader_vendor_note, poheader_over_total, poheader_frequency_gid, poheader_from_month, poheader_to_month, poheader_type, poheader_status, "
                lsSQL &= " poheader_termcond_gid, poheader_add_termandcond, poheader_currentapprovalstage, poheader_projectmanager, poheader_migstatus )  "

                lsSQL &= " VALUES ('" & lsPONumber & "'," & lsPODate & "," & lsPOEndDate & "," & liRaisorGID & ",'" & Mid(lsPOITType, 1, 1) & "',"
                lsSQL &= " '" & IIf(Mid(lsPORequestFor.Trim, 1, 1) = "P", 3, 4) & "'," & liVendorGID & ",'" & lsPOVendorNote & "'," & Val(lsPOOverTotal) & ",1,"
                lsSQL &= "'JAN','DEC','" & lsPOType & "',5,0,'" & lsPOAdditionalTerm & "',Null," & liProjectOwnerGID & ",'" & Mid(lsStatus, 1, 1) & "')"

                lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)

                lsSQL = ""
                lsSQL = " SELECT poheader_gid FROM fb_trn_tpoheader "
                lsSQL &= " WHERE poheader_pono='" & lsPONumber & "' "

                liPOGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            End If

            If liPOGID <> 0 Then
                ' INSERTING DATA INTO PO DETAILS
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL &= " INSERT INTO fb_trn_tpodetails ( podetails_pohead_gid, podetails_prodservice_gid, podetails_desc, podetails_serv_month, podetails_unitprice, "
                lsSQL &= " podetails_total, podetails_cbfdet_gid, podetails_qty, podetails_uom_gid, podetails_sno) VALUES(" & liPOGID & "," & liProductGID & ",'" & lsDescription & "','JAN-DEC',"
                lsSQL &= lsUnitPrice & "," & lsTotal & "," & liCBFDetailsGID & "," & lsQty & "," & liUOMGID & "," & lsSLNO & ")"

                lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)
            Else
                lsErrNote &= " PO# NOT FOUND : "
            End If

NextFetch:
            If lsErrMessage = "" And lsErrNote = "" Then
                liValid += 1
            Else
                liError += 1

                lobjErrorDatatable.Rows.Add()

                With lobjErrorDatatable.Rows(lobjErrorDatatable.Rows.Count - 1)
                    .Item("SlNO") = lsSLNO
                    .Item("poheader_pono") = lsPONumber
                    .Item("podetails_cbfheader_cbfno") = lsPOCBFNumber
                    .Item("CBFDetail_Sno") = lsPOCBFSLNumber

                    .Item("poheader_date") = lsPODate
                    .Item("poheader_enddate") = lsPOEndDate
                    .Item("poheader_raisor") = lsPORaisor
                    .Item("poheader_projectmanager") = lsPOProjectManager

                    .Item("poheader_requestfor") = lsPORequestFor
                    .Item("poheader_ittype") = lsPOITType
                    .Item("poheader_vendor") = lsPOVendor
                    .Item("poheader_vendor_note") = lsPOVendorNote

                    .Item("poheader_over_total") = lsPOOverTotal
                    .Item("poheader_type") = lsPOType
                    .Item("poheader_termcond") = lsPOTermCond
                    .Item("poheader_add_termandcond") = lsPOAdditionalTerm

                    .Item("podetails_prodservice") = lsProduct_service
                    .Item("podetails_desc") = lsDescription
                    .Item("podetails_uom") = lsUOM
                    .Item("podetails_qty") = lsQty

                    .Item("podetails_unitprice") = lsUnitPrice
                    .Item("podetails_discount") = lsDiscount
                    .Item("podetails_base_amt") = lsBaseAmount
                    .Item("podetails_tax1") = lsTax1

                    .Item("podetails_tax2") = lsTax2
                    .Item("podetails_tax3") = lsTax3
                    .Item("podetails_total") = lsTotal

                    .Item("Error") = lsErrMessage
                    .Item("Error Note") = lsErrNote
                End With
            End If

            txtStatus.Text = liValid & " Of " & lobjDataTable.Rows.Count & " Records Migrated...  Error - " & liError
            Application.DoEvents()
        Next

        If lobjErrorDatatable.Rows.Count > 0 Then
            PrintDGridviewXML(lobjErrorDatatable, Application.StartupPath & "\PO-MIG-ERR.xls")
            MessageBox.Show("Descrepance Records Spooled @ " & Application.StartupPath & "\PO-MIG-ERR.xls")
        Else
            MessageBox.Show("Imported successfully...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        'Catch ex As Exception
        '    grpMain.Enabled = True
        '    MsgBox(ex.Message, MsgBoxStyle.Critical)
        'End Try
    End Sub


    Public Sub PO_Shipment_Data_Migration()
        'Try
        Dim lsSQL As String = ""
        Dim lsFileName As String = ""
        Dim lsResult As String = ""
        Dim liDumpGid As Integer = 0
        Dim lobjOleDbConnection As New OleDb.OleDbConnection
        Dim lobjDataAdapter As OleDb.OleDbDataAdapter
        Dim lobjDataTable As DataTable
        Dim lobjDataSet As New DataSet
        Dim liError As Integer
        Dim liValid As Integer

        Dim lsErrMessage As String = ""
        Dim lsErrNote As String = ""

        lsFileName = Mid(txtFileName.Text.Trim, InStrRev(txtFileName.Text.Trim, "\") + 1)

        Try
            If lobjOleDbConnection.State = 1 Then lobjOleDbConnection.Close()
            With lobjOleDbConnection

                If Microsoft.VisualBasic.Right(lsFileName, 4) = "xlsx" Then
                    'read a 2007 file   
                    .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                                         txtFileName.Text.Trim & ";" + ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
                Else
                    'read a 97-2003 file   
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" & _
                                            txtFileName.Text.Trim & ";" + "Extended Properties=Excel 8.0;"
                End If

                .Open()

            End With

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            lobjOleDbConnection.Close()
            Exit Sub
        End Try

        lsSQL = "SELECT * FROM [" & cmbSheet.Text & "$]"

        lobjDataAdapter = New OleDb.OleDbDataAdapter(lsSQL, lobjOleDbConnection)
        lobjDataAdapter.Fill(lobjDataSet, "DATA")
        lobjDataTable = lobjDataSet.Tables("DATA")
        lobjOleDbConnection.Close()

        Dim liIndex As Integer

        For liIndex = 0 To lobjDataTable.Columns.Count - 1
            If Not InStr(fsColumnHeaders.ToLower.Trim, "," & lobjDataTable.Columns(liIndex).ColumnName.ToString.Trim.ToLower & ",") > 0 Then
                MessageBox.Show("PO Shipment Dump is not in the correct format...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Next

        'SlNO	poheader_pono	poshipment_shipmenttype	poshipment_branch	poshipment_remarks	podetails_qty	poshipment_incharge

        Dim lsSLNO As String
        Dim lsPONumber As String
        Dim lsShipmentType As String
        Dim lsBranch As String

        Dim lsRemarks As String
        Dim lsQty As String

        Dim lsIncharge As String
        Dim lsStatus As String


        Dim liPOGID As Integer
        Dim liPODETGID As Integer

        Dim liInchargeGID As Integer
        Dim liBranchGID As Integer

        Dim liPOShipmentGID As Integer
        Dim liGRNReleaseforPOGID As Integer
        Dim liGRNInwardGID As Integer
        Dim liMX As Integer

        Dim ldTotal As Double


        Dim lobjErrorDatatable As New DataTable

        With lobjErrorDatatable
            .Columns.Add("SlNO")
            .Columns.Add("poheader_pono")
            .Columns.Add("poshipment_shipmenttype")
            .Columns.Add("poshipment_branch")

            .Columns.Add("poshipment_remarks")
            .Columns.Add("podetails_qty")
            .Columns.Add("poshipment_incharge")

            .Columns.Add("Error")
            .Columns.Add("Error Note")
        End With

        For i As Integer = 0 To lobjDataTable.Rows.Count - 1
            lsErrMessage = ""
            lsErrNote = ""
            With lobjDataTable.Rows(i)
                lsSLNO = QuoteFilter(.Item("SlNO").ToString)
                lsPONumber = QuoteFilter(.Item("poheader_pono").ToString)
                lsShipmentType = QuoteFilter(.Item("poshipment_shipmenttype").ToString)
                lsBranch = QuoteFilter(.Item("poshipment_branch").ToString)
                lsRemarks = QuoteFilter(.Item("poshipment_remarks").ToString)
                lsQty = QuoteFilter(.Item("podetails_qty").ToString)
                lsIncharge = QuoteFilter(.Item("poshipment_incharge").ToString)
            End With

            liInchargeGID = 0
            liPOGID = 0
            liPODETGID = 0


            ' EMPLOYEE GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT employee_gid FROM iem_mst_temployee "
            lsSQL &= " WHERE employee_code='" & Mid(lsIncharge, 1, InStr(lsIncharge, "-") - 1).Trim & "' "

            liInchargeGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            If InStr(lsBranch, "-") = 0 Then lsBranch &= " - "

            ' BRANCH GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT branch_gid FROM iem_mst_tbranch "
            lsSQL &= " WHERE branch_code='" & Mid(lsBranch, 1, InStr(lsBranch, "-") - 1).Trim & "' "

            liBranchGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            ' PO HEADER GID FINDING 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT poheader_gid FROM fb_trn_tpoheader "
            lsSQL &= " WHERE poheader_pono='" & lsPONumber & "' "

            liPOGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            ' PO MIGRATION STATUS
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT poheader_migstatus FROM fb_trn_tpoheader "
            lsSQL &= " WHERE poheader_pono='" & lsPONumber & "' "

            lsStatus = loDBConnection.GetExecuteScalar(lsSQL).ToString


            ' PO DETAILS GID FINDING 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT podetails_gid FROM fb_trn_tpodetails "
            lsSQL &= " WHERE podetails_pohead_gid=" & liPOGID
            lsSQL &= " AND podetails_sno= " & Val(lsSLNO)

            liPODETGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            lsSQL = ""
            lsSQL = " SELECT podetails_total FROM fb_trn_tpodetails "
            lsSQL &= " WHERE podetails_pohead_gid=" & liPOGID
            lsSQL &= " AND podetails_sno= " & Val(lsSLNO)

            ldTotal = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            'Select Case 0
            If liPOGID = 0 Then _
                lsErrNote &= " PO# NOT FOUND : "

            If liPODETGID = 0 Then _
                lsErrNote &= " PO ITEM NOT FOUND : "

            If liInchargeGID = 0 Then _
                lsErrNote &= " INCHARGE NOT FOUND : "

            If liBranchGID = 0 Then _
                lsErrNote &= " BRANCH NOT FOUND : "
            'End Select

            If lsErrNote <> "" Then GoTo NextFetch

            lsSQL = ""
            lsSQL &= " INSERT INTO fb_trn_tposhipment ( "
            lsSQL &= " poshipment_podet_gid,poshipment_type_gid,poshipment_branch_gid,"
            lsSQL &= " poshipment_qty,poshipment_isremoved,poshipment_empgid,"
            lsSQL &= " poshipment_isamended,poshipment_remarks) VALUES("
            lsSQL &= liPODETGID & ",1," & liBranchGID & "," & lsQty & ",'N'," & liInchargeGID & ",'N','" & lsRemarks & "') "

            lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)


            If lsStatus <> "C" Then GoTo NextFetch

            lsSQL = ""
            lsSQL = " SELECT poshipment_gid FROM fb_trn_tposhipment "
            lsSQL &= " WHERE poshipment_podet_gid=" & liPODETGID

            liPOShipmentGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            If liPOShipmentGID <> 0 Then

                '' INSERTING DATA INTO PO SHIPMENT DETAILS
                ''-------------------------------------------------------------------------------------------------------------------------------
                'lsSQL = ""
                'lsSQL &= " INSERT INTO fb_trn_tposhipment ( "
                'lsSQL &= " poshipment_podet_gid,poshipment_type_gid,poshipment_branch_gid,"
                'lsSQL &= " poshipment_qty,poshipment_isremoved,poshipment_empgid,"
                'lsSQL &= " poshipment_isamended,poshipment_remarks) VALUES("
                'lsSQL &= liPODETGID & ",1," & liBranchGID & "," & lsQty & ",'N'," & liRaisorGID & ",'N','Test Wo Migration') "

                'lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)


                ' INSERTING DATA INTO GRN RELEASE FOR PO
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL &= "INSERT INTO fb_trn_tgrnreleaseforpo ("
                lsSQL &= " grnreleaseforpo_podet_gid,grnreleaseforpo_released_qty,grnreleaseforpo_isremoved,"
                lsSQL &= " grnreleaseforpo_balanceqty,grnreleaseforpo_poshipment_gid,grnreleaseforpo_branch_type,"
                lsSQL &= " grnreleaseforpo_released_date,grnreleaseforpo_releasedby,inward_release) VALUES("
                lsSQL &= liPODETGID & "," & lsQty & ",'N',0," & liPOShipmentGID & ",'D','" & Format(Now, "yyyy-MM-dd") & "', "
                lsSQL &= liInchargeGID & ",Null)"

                lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)

                lsSQL = ""
                lsSQL = " SELECT grnreleaseforpo_gid FROM fb_trn_tgrnreleaseforpo "
                lsSQL &= " WHERE grnreleaseforpo_podet_gid=" & liPODETGID
                lsSQL &= " AND grnreleaseforpo_poshipment_gid=" & liPOShipmentGID

                liGRNReleaseforPOGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


                ' INSERTING DATA INTO GRN INWARD
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL &= " INSERT INTO fb_trn_tgrninwrdheader(grninwrdheader_refno,grninwrdheader_grndatetime,grninwardheader_poheader,grninwrdheader_rasiergid,"
                lsSQL &= " grninwrdheader_dcno,grninwrdheader_invoiceno,grninwrdheader_remarks,grninwrdheader_isremoved,grninwrdheader_status,grninwrdheader_branch_type,"
                lsSQL &= " grniwrdheader_capstatus) VALUES("
                lsSQL &= "'GRN" & Format(Now, "yyMMdd") & Format(i, "0000") & "',SYSDATETIME()," & liPOGID & "," & liInchargeGID & ","
                lsSQL &= "'12345','12345','PO DATA Migration','N','5','D',0) "

                lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)

                lsSQL = ""
                lsSQL = " SELECT grninwrdheader_gid FROM fb_trn_tgrninwrdheader "
                lsSQL &= " WHERE grninwardheader_poheader=" & liPOGID

                liGRNInwardGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


                ' INSERTING DATA INTO GRN INWARD DETAILS 
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL &= " INSERT INTO fb_trn_tgrninwrddet( "
                lsSQL &= " grninwrddet_grnreleforpo_gid,grninwrddet_reced_qty,grninwrddet_reced_date,"
                lsSQL &= " grninwrddet_isremoved,grninwrddet_grninwrdhead_gid,grninwrddet_assetsrlno, "
                lsSQL &= " grninwrddet_puttousedatetime,grninwrddet_mft_name,grninwrddet_capstatus ) VALUES( "
                lsSQL &= liGRNReleaseforPOGID & "," & lsQty & ",SYSDATETIME(),'N'," & liGRNInwardGID & ",'',SYSDATETIME(),'',0) "

                lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)


                ' INSERTING DATA INTO GRN CONFIRMATION 
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL &= " INSERT INTO fb_trn_tgrnconfrm(grnconfrm_grninwrdheader_gid,grnconfrm_remarks,grnconfrm_date,grnconfrm_confirmby,grnconfrm_status,"
                lsSQL &= " grnconfrm_isremoved)values(" & liGRNInwardGID & ",'Test Migration',SYSDATETIME() ," & liInchargeGID & ",'5','N')"

                lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)


                ' INSERTING DATA INTO ECF
                '-------------------------------------------------------------------------------------------------------------------------------

                ' INSERTING DATA INTO INVOICE
                '-------------------------------------------------------------------------------------------------------------------------------

                ' INSERTING DATA INTO INVOICE PO
                '-------------------------------------------------------------------------------------------------------------------------------

                lsSQL = ""
                lsSQL &= " INSERT INTO iem_trn_tinvoicepo (invoicepo_invoice_gid,invoicepo_po_gid,invoicepo_mapped_amount,invoicepo_isremoved) "
                lsSQL &= " VALUES(0," & liPOGID & "," & ldTotal & ",'N') "

                lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)

                lsSQL = ""
                lsSQL = " SELECT MAX(invoicepo_gid) FROM iem_trn_tinvoicepo "
                'lsSQL &= " WHERE poshipment_podet_gid=" & liPODETGID

                liMX = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


                ' INSERTING DATA INTO INVOICE PO ITEM 
                '-------------------------------------------------------------------------------------------------------------------------------

                lsSQL = ""
                lsSQL &= " INSERT INTO iem_trn_tinvoicepoitem (invoicepoitem_po_gid,invoicepoitem_invoice_gid,invoicepoitem_poitem_gid,"
                lsSQL &= " invoicepoitem_qty, invoicepoitem_rate,invoicepoitem_amount,invoicepoitem_isremoved)  "
                lsSQL &= " VALUES(" & liMX & ",0," & liPODETGID & "," & lsQty & ",0," & ldTotal & ",'N') "

                lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)

            Else
                lsErrNote &= " SHIPMENT ID NOT GENERATED : "
            End If

NextFetch:
            If lsErrMessage = "" And lsErrNote = "" Then
                liValid += 1
            Else
                liError += 1

                lobjErrorDatatable.Rows.Add()

                With lobjErrorDatatable.Rows(lobjErrorDatatable.Rows.Count - 1)
                    .Item("SlNO") = lsSLNO
                    .Item("poheader_pono") = lsPONumber
                    .Item("poshipment_shipmenttype") = lsShipmentType
                    .Item("poshipment_branch") = lsBranch
                    .Item("poshipment_remarks") = lsRemarks
                    .Item("podetails_qty") = lsQty
                    .Item("poshipment_incharge") = lsIncharge

                    .Item("Error") = lsErrMessage
                    .Item("Error Note") = lsErrNote

                End With
            End If

            txtStatus.Text = liValid & " Of " & lobjDataTable.Rows.Count & " Records Migrated...  Error - " & liError
            Application.DoEvents()
        Next

        If lobjErrorDatatable.Rows.Count > 0 Then
            PrintDGridviewXML(lobjErrorDatatable, Application.StartupPath & "\POSHIP-MIG-ERR.xls")
            MessageBox.Show("Descrepance Records Spooled @ " & Application.StartupPath & "\POSHIP-MIG-ERR.xls")
        Else
            MessageBox.Show("Imported successfully...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        'Catch ex As Exception
        '    grpMain.Enabled = True
        '    MsgBox(ex.Message, MsgBoxStyle.Critical)
        'End Try
    End Sub


    Public Sub CBF_Data_Migration()
        'Try
        Dim lsSQL As String = ""
        Dim lsFileName As String = ""
        Dim lsResult As String = ""
        Dim liDumpGid As Integer = 0
        Dim lobjOleDbConnection As New OleDb.OleDbConnection
        Dim lobjDataAdapter As OleDb.OleDbDataAdapter
        Dim lobjDataTable As DataTable
        Dim lobjDataSet As New DataSet
        Dim liError As Integer
        Dim liValid As Integer

        Dim lsErrMessage As String = ""
        Dim lsErrNote As String = ""


        lsFileName = Mid(txtFileName.Text.Trim, InStrRev(txtFileName.Text.Trim, "\") + 1)

        Try
            If lobjOleDbConnection.State = 1 Then lobjOleDbConnection.Close()
            With lobjOleDbConnection

                If Microsoft.VisualBasic.Right(lsFileName, 4) = "xlsx" Then
                    'read a 2007 file   
                    .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                                         txtFileName.Text.Trim & ";" + ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
                Else
                    'read a 97-2003 file   
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" & _
                                            txtFileName.Text.Trim & ";" + "Extended Properties=Excel 8.0;"
                End If

                .Open()

            End With

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            lobjOleDbConnection.Close()
            Exit Sub
        End Try

        lsSQL = "SELECT * FROM [" & cmbSheet.Text & "$]"

        lobjDataAdapter = New OleDb.OleDbDataAdapter(lsSQL, lobjOleDbConnection)
        lobjDataAdapter.Fill(lobjDataSet, "DATA")
        lobjDataTable = lobjDataSet.Tables("DATA")
        lobjOleDbConnection.Close()

        Dim liIndex As Integer

        For liIndex = 0 To lobjDataTable.Columns.Count - 1
            If Not InStr(fsColumnHeaders.ToLower.Trim, "," & lobjDataTable.Columns(liIndex).ColumnName.ToString.Trim.ToLower & ",") > 0 Then
                MessageBox.Show("CBF Dump is not in the correct format...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Next

        Dim lsCBFNumber As String
        Dim lsDetailSno As String
        Dim lsCBFOBFFlag As String
        Dim lsStartDate As String
        Dim lsEndDate As String
        Dim lsProjectOwner As String
        Dim lsBranchCode As String
        Dim lsMode As String
        Dim lsApprovalType As String
        Dim lsIsBudgeted As String
        Dim lsDeviationAmount As String
        Dim lsCBFAmount As String
        Dim lsDescription As String
        Dim lsRaiser As String
        Dim lsRequestFor As String
        Dim lsRemarks As String
        Dim lsIsBranchSingle As String
        Dim lsBudgetOwner As String
        Dim lsPARPRDescription As String
        Dim lsProductService As String
        Dim lsCBFDetailsDescription As String
        Dim lsUOM As String
        Dim lsQty As String
        Dim lsUnitPrice As String
        Dim lsTotalAmount As String
        Dim lsCBFDetailsRemarks As String
        Dim lsCOA As String
        Dim lsFCCC As String
        Dim lsBudgetLine As String
        Dim lsVendor As String
        Dim lsProductGroup As String

        Dim liCBFGID As Integer
        Dim liProductGID As Integer
        Dim liUOMGID As Integer
        Dim liVendorGID As Integer
        Dim liProdServiceGID As Integer
        Dim liCBFDetailsGID As Integer

        Dim liRaisorGID As Integer
        Dim liProjectOwnerGID As Integer
        Dim liBudgetOwnerGId As Integer

        Dim liBranchGID As Integer

        Dim lobjErrorDatatable As New DataTable


        With lobjErrorDatatable
            .Columns.Add("CBFNo")
            .Columns.Add("Detail_Sno")
            .Columns.Add("CBF_OBF_Flag")
            .Columns.Add("Start_Date")
            .Columns.Add("End_Date")
            .Columns.Add("Project_Owner")
            .Columns.Add("Branch")
            .Columns.Add("Mode")
            .Columns.Add("Approval_Type")
            .Columns.Add("Is_Budgeted")
            .Columns.Add("Deviation_Amount")
            .Columns.Add("CBF_Amount")

            .Columns.Add("Description")
            .Columns.Add("Raiser")
            .Columns.Add("Request_For")
            .Columns.Add("Remarks")
            .Columns.Add("Is_Branch_Single")
            .Columns.Add("Budget_Owner")
            .Columns.Add("PAR_PR_Description")
            .Columns.Add("Product_Service")
            .Columns.Add("CBF_Details_Description")
            .Columns.Add("UOM")
            .Columns.Add("QTY")
            .Columns.Add("Unit_Price")

            .Columns.Add("Total_Amount")
            .Columns.Add("CBF_Details_Remarks")
            .Columns.Add("COA")
            .Columns.Add("FCCC")
            .Columns.Add("Budget_Line")
            .Columns.Add("Vendor")
            .Columns.Add("Product_Group")

            .Columns.Add("Error")
            .Columns.Add("Error Note")
        End With

        For i As Integer = 0 To lobjDataTable.Rows.Count - 1
            lsErrMessage = ""
            lsErrNote = ""

            With lobjDataTable.Rows(i)

                lsCBFNumber = QuoteFilter(.Item("CBFNo").ToString)
                lsDetailSno = QuoteFilter(.Item("Detail_Sno").ToString)
                lsCBFOBFFlag = QuoteFilter(.Item("CBF_OBF_Flag").ToString)
                lsStartDate = QuoteFilter(.Item("Start_Date").ToString)
                lsEndDate = QuoteFilter(.Item("End_Date").ToString)
                lsProjectOwner = QuoteFilter(.Item("Project_Owner").ToString)
                lsBranchCode = QuoteFilter(.Item("Branch").ToString)
                lsMode = QuoteFilter(.Item("Mode").ToString)
                lsApprovalType = QuoteFilter(.Item("Approval_Type").ToString)
                lsIsBudgeted = QuoteFilter(.Item("Is_Budgeted").ToString)
                lsDeviationAmount = QuoteFilter(.Item("Deviation_Amount").ToString)
                lsCBFAmount = QuoteFilter(.Item("CBF_Amount").ToString)

                lsDescription = QuoteFilter(.Item("Description").ToString)
                lsRaiser = QuoteFilter(.Item("Raiser").ToString)
                lsRequestFor = QuoteFilter(.Item("Request_For").ToString)
                lsRemarks = QuoteFilter(.Item("Remarks").ToString)
                lsIsBranchSingle = QuoteFilter(.Item("Is_Branch_Single").ToString)
                lsBudgetOwner = QuoteFilter(.Item("Budget_Owner").ToString)
                lsPARPRDescription = QuoteFilter(.Item("PAR_PR_Description").ToString)
                lsProductService = QuoteFilter(.Item("Product_Service").ToString)
                lsCBFDetailsDescription = QuoteFilter(.Item("CBF_Details_Description").ToString)
                lsUOM = QuoteFilter(.Item("UOM").ToString)
                lsQty = QuoteFilter(.Item("QTY").ToString)
                lsUnitPrice = QuoteFilter(.Item("Unit_Price").ToString)

                lsTotalAmount = QuoteFilter(.Item("Total_Amount").ToString)
                lsCBFDetailsRemarks = QuoteFilter(.Item("CBF_Details_Remarks").ToString)
                lsCOA = QuoteFilter(.Item("Chart_Of_Acc").ToString)
                lsFCCC = QuoteFilter(.Item("FCCC").ToString)
                lsBudgetLine = QuoteFilter(.Item("Budget_Line").ToString)
                lsVendor = QuoteFilter(.Item("Vendor").ToString)
                lsProductGroup = QuoteFilter(.Item("Product_Group").ToString)
            End With

            If lsStartDate <> "" Then
                If IsDate(lsStartDate) Then
                    lsStartDate = "'" & Format(CDate(lsStartDate), "yyyy-MM-dd") & "'"
                Else
                    lsStartDate = "NULL"
                End If
            Else
                lsStartDate = "NULL"
            End If

            If lsEndDate <> "" Then
                If IsDate(lsEndDate) Then
                    lsEndDate = "'" & Format(CDate(lsEndDate), "yyyy-MM-dd") & "'"
                Else
                    lsEndDate = "NULL"
                End If
            Else
                lsEndDate = "NULL"
            End If


            If lsFCCC.Trim = "NA" Then lsFCCC = "11103"


            liCBFGID = 0
            liProductGID = 0
            liUOMGID = 0
            liVendorGID = 0
            liProdServiceGID = 0
            liCBFDetailsGID = 0
            liRaisorGID = 0
            liProjectOwnerGID = 0
            liBudgetOwnerGId = 0
            liBranchGID = 0


            ' UOM GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT uom_gid FROM iem_mst_tuom "
            lsSQL &= " WHERE LOWER(uom_code)='" & lsUOM.ToLower.Trim & "'"

            liUOMGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            ' PRODUCT / SERVICE GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            If InStr(lsProductService, "-") = 0 Then _
                lsProductService &= " - "

            lsSQL = ""
            lsSQL = " SELECT prodservice_gid FROM fb_mst_tprodservice "
            lsSQL &= " WHERE prodservice_code='" & Mid(lsProductService, 1, InStr(lsProductService, "-") - 1).Trim & "' "

            liProductGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            lsSQL = ""
            lsSQL = " SELECT prodservice_prodservicegid FROM fb_mst_tprodservice "
            lsSQL &= " WHERE prodservice_code='" & Mid(lsProductService, 1, InStr(lsProductService, "-") - 1).Trim & "' "

            liProdServiceGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            ' SUPPLIER GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT supplierheader_gid FROM asms_trn_tsupplierheader "
            lsSQL &= " WHERE supplierheader_name='" & lsVendor & "' "

            liVendorGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            ' EMPLOYEE GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT employee_gid FROM iem_mst_temployee "
            lsSQL &= " WHERE employee_code='" & Mid(lsRaiser, 1, InStr(lsRaiser, "-") - 1).Trim & "' "

            liRaisorGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            ' PROJECT OWNER GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT employee_gid FROM iem_mst_temployee "
            lsSQL &= " WHERE employee_code='" & Mid(lsProjectOwner, 1, InStr(lsProjectOwner, "-") - 1).Trim & "' "

            liProjectOwnerGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            lsSQL = ""
            lsSQL = " SELECT projectowner_gid FROM iem_mst_tprojectowner "
            lsSQL &= " WHERE projectowner_employeegid=" & liProjectOwnerGID & " "

            liProjectOwnerGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            ' BUDGET OWNER GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT employee_gid FROM iem_mst_temployee "
            lsSQL &= " WHERE employee_code='" & Mid(lsBudgetOwner, 1, InStr(lsBudgetOwner, "-") - 1).Trim & "' "

            liBudgetOwnerGId = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            ' BRANCH GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT branch_gid FROM iem_mst_tbranch "
            lsSQL &= " WHERE branch_code='" & Mid(lsBranchCode, 1, InStr(lsBranchCode, "-") - 1).Trim & "' "

            liBranchGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)



            If liProductGID = 0 Then _
                    lsErrNote &= " PRODUCT NOT FOUND : "

            If liProdServiceGID = 0 Then _
                lsErrNote &= " PRODUCT SERVICE NOT FOUND : "

            If liVendorGID = 0 Then _
                lsErrNote &= " VENDOR NOT FOUND : "

            If liRaisorGID = 0 Then _
                lsErrNote &= " RAISER NOT FOUND : "

            If liProjectOwnerGID = 0 Then _
                lsErrNote &= " PROJECT OWNER NOT FOUND : "

            If liBudgetOwnerGId = 0 Then _
                lsErrNote &= " BUDGET OWNER NOT FOUND : "

            If liBranchGID = 0 Then _
                lsErrNote &= " BRANCH NOT FOUND : "

            If lsErrNote <> "" Then GoTo NextFetch

            ' CBF HEADER GID FINDING 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT cbfheader_gid FROM fb_trn_tcbfheader "
            lsSQL &= " WHERE cbfheader_cbfno='" & lsCBFNumber & "' "

            liCBFGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            If liCBFGID = 0 Then

                ' INSERTING DATA INTO CBF HEADER 
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL &= " INSERT INTO fb_trn_tcbfheader(cbfheader_cbfno,cbfheader_cbfobf_flag,cbfheader_date,cbfheader_enddate,cbfheader_projectowner,cbfheader_branch_gid,"
                lsSQL &= " cbfheader_mode,cbfheader_prpar_gid,cbfheader_approvaltype,cbfheader_isbudgeted,cbfheader_Devi_amt,cbfheader_cbfamt,cbfheader_desc,"
                lsSQL &= " cbfheader_rasier_gid,cbfheader_requestfor_gid,cbfheader_requesttype,cbfheader_remarks,cbfheader_budgetowner_gid,cbfheader_status)  "

                lsSQL &= " VALUES('" & lsCBFNumber & "','" & Mid(lsCBFOBFFlag.Trim, 1, 1) & "'," & lsStartDate & "," & lsEndDate & "," & liProjectOwnerGID & ","
                lsSQL &= liBranchGID & ",'" & lsMode & "',0,'" & Mid(lsApprovalType, 1, 1) & "','" & lsIsBudgeted & "'," & lsDeviationAmount & "," & lsCBFAmount & ","
                lsSQL &= "'" & lsDescription & "'," & liRaisorGID & "," & IIf(lsRequestFor.Trim = "IT", 3, 4) & ",'','" & lsRemarks.Trim & "'," & liBudgetOwnerGId & ",5) "

                lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)

                lsSQL = ""
                lsSQL = " SELECT cbfheader_gid FROM fb_trn_tcbfheader "
                lsSQL &= " WHERE cbfheader_cbfno='" & lsCBFNumber & "' "

                liCBFGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            End If

            If liCBFGID <> 0 Then

                lsSQL = ""
                lsSQL = " SELECT cbfdetails_gid FROM fb_trn_tcbfdetails "
                lsSQL &= " WHERE cbfdetails_cbfhead_gid='" & liCBFGID & "' AND cbfdetails_sno = " & lsDetailSno

                liCBFDetailsGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

                If liCBFDetailsGID <> 0 Then
                    lsSQL = ""
                    lsSQL &= " DELETE FROM fb_trn_tcbfdetails "
                    lsSQL &= " WHERE cbfdetails_cbfhead_gid='" & liCBFGID & "' AND cbfdetails_sno = " & lsDetailSno

                    loDBConnection.ExecuteNonQuerySQL(lsSQL)
                End If

                ' INSERTING DATA INTO CBF DETAILS 
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL &= " INSERT INTO fb_trn_tcbfdetails(cbfheader_cbfobf_flag,cbfdetails_cbfhead_gid,cbfdetails_parprdesc,cbfdetails_year,cbfdetails_prod_gid,"
                lsSQL &= " cbfdetails_desc,cbfdetails_uom_gid,cbfdetails_qty,cbfdetails_unitprice,cbfdetails_totalamt,cbfdetails_remarks,cbfdetails_chartofacc,cbfdetails_fccc,"
                lsSQL &= " cbfdetails_budgetline,cbfdetails_budgetowner_gid,cbfdetails_vendor_gid,cbfdetails_prpardel_gid,cbfdetails_prodservgrp_gid,"
                lsSQL &= " cbfdetails_sno) "

                lsSQL &= " VALUES ('" & Mid(lsCBFOBFFlag.Trim, 1, 1) & "'," & liCBFGID & ",'" & lsPARPRDescription & "',''," & liProductGID & ",'" & lsCBFDetailsDescription & "'," & liUOMGID & "," & Val(lsQty)
                lsSQL &= "," & Val(lsUnitPrice) & "," & Val(lsTotalAmount) & ",'" & lsCBFDetailsRemarks & "','" & lsCOA & "','" & Mid(lsFCCC.Trim, 1, 5) & "'," & lsBudgetLine & "," & liBudgetOwnerGId & "," & liVendorGID & ",0," & liProdServiceGID
                lsSQL &= "," & lsDetailSno & ")"

                lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)
            Else
                lsErrNote &= " CBF NOT FOUND "
            End If

NextFetch:
            If lsErrMessage = "" And lsErrNote = "" Then
                liValid += 1
            Else
                liError += 1

                lobjErrorDatatable.Rows.Add()

                With lobjErrorDatatable.Rows(lobjErrorDatatable.Rows.Count - 1)

                    .Item("CBFNo") = lsCBFNumber
                    .Item("Detail_Sno") = lsDetailSno
                    .Item("CBF_OBF_Flag") = lsCBFOBFFlag
                    .Item("Start_Date") = lsStartDate
                    .Item("End_Date") = lsEndDate
                    .Item("Project_Owner") = lsProjectOwner
                    .Item("Branch") = lsBranchCode
                    .Item("Mode") = lsMode
                    .Item("Approval_Type") = lsApprovalType
                    .Item("Is_Budgeted") = lsIsBudgeted
                    .Item("Deviation_Amount") = lsDeviationAmount
                    .Item("CBF_Amount") = lsCBFAmount

                    .Item("Description") = lsDescription
                    .Item("Raiser") = lsRaiser
                    .Item("Request_For") = lsRequestFor
                    .Item("Remarks") = lsRemarks
                    .Item("Is_Branch_Single") = lsIsBranchSingle
                    .Item("Budget_Owner") = lsBudgetOwner
                    .Item("PAR_PR_Description") = lsPARPRDescription
                    .Item("Product_Service") = lsProductService
                    .Item("CBF_Details_Description") = lsCBFDetailsDescription
                    .Item("UOM") = lsUOM
                    .Item("QTY") = lsQty
                    .Item("Unit_Price") = lsUnitPrice

                    .Item("Total_Amount") = lsTotalAmount
                    .Item("CBF_Details_Remarks") = lsCBFDetailsRemarks
                    .Item("COA") = lsCOA
                    .Item("FCCC") = lsFCCC
                    .Item("Budget_Line") = lsBudgetLine
                    .Item("Vendor") = lsVendor
                    .Item("Product_Group") = lsProductGroup

                    .Item("Error") = lsErrMessage
                    .Item("Error Note") = lsErrNote
                End With
            End If

            txtStatus.Text = liValid & " Of " & lobjDataTable.Rows.Count & " Records Migrated...  Error - " & liError
            Application.DoEvents()
        Next

        If lobjErrorDatatable.Rows.Count > 0 Then
            PrintDGridviewXML(lobjErrorDatatable, Application.StartupPath & "\CBF-MIG-ERR.xls")
            MessageBox.Show("Descrepance Records Spooled @ " & Application.StartupPath & "\CBF-MIG-ERR.xls")
        Else
            MessageBox.Show("Imported successfully...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If


        'Catch ex As Exception
        '    grpMain.Enabled = True
        '    MsgBox(ex.Message, MsgBoxStyle.Critical)
        'End Try
    End Sub

    Public Sub ECF_Data_Migration()
        'Try
        Dim lsSQL As String = ""
        Dim lsFileName As String = ""
        Dim lsResult As String = ""
        Dim liDumpGid As Integer = 0
        Dim lobjOleDbConnection As New OleDb.OleDbConnection
        Dim lobjDataAdapter As OleDb.OleDbDataAdapter
        Dim lobjDataTable As DataTable
        Dim lobjDataSet As New DataSet
        Dim liError As Integer
        Dim liValid As Integer

        Dim lsErrMessage As String = ""
        Dim lsErrNote As String = ""


        lsFileName = Mid(txtFileName.Text.Trim, InStrRev(txtFileName.Text.Trim, "\") + 1)

        Try
            If lobjOleDbConnection.State = 1 Then lobjOleDbConnection.Close()
            With lobjOleDbConnection

                If Microsoft.VisualBasic.Right(lsFileName, 4) = "xlsx" Then
                    'read a 2007 file   
                    .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                                         txtFileName.Text.Trim & ";" + ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
                Else
                    'read a 97-2003 file   
                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" & _
                                            txtFileName.Text.Trim & ";" + "Extended Properties=Excel 8.0;"
                End If

                .Open()

            End With

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            lobjOleDbConnection.Close()
            Exit Sub
        End Try

        lsSQL = "SELECT * FROM [" & cmbSheet.Text & "$]"

        lobjDataAdapter = New OleDb.OleDbDataAdapter(lsSQL, lobjOleDbConnection)
        lobjDataAdapter.Fill(lobjDataSet, "DATA")
        lobjDataTable = lobjDataSet.Tables("DATA")
        lobjOleDbConnection.Close()

        Dim liIndex As Integer

        For liIndex = 0 To lobjDataTable.Columns.Count - 1
            If Not InStr(fsColumnHeaders.ToLower.Trim, "," & lobjDataTable.Columns(liIndex).ColumnName.ToString.Trim.ToLower & ",") > 0 Then
                MessageBox.Show("ECF Dump is not in the correct format...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If
        Next


        Dim lsRowID As String
        Dim lsBrCode As String
        Dim lsAuthDt As String
        Dim lsSFBatchNo As String

        Dim lsSFDocNumb As String
        Dim lsSFNumbil As String
        Dim lsSFBillNumb As String
        Dim lsSFVendCode As String

        Dim lsSFVenName As String
        Dim lsSFActualVendor As String
        Dim lsSFBillDate As String
        Dim lsSFBillStDt As String
        Dim lsSFBillEDDt As String

        Dim lsSFBillAmount As String
        Dim lsSFBillDesc As String
        Dim lsSFNetAmount As String
        Dim lsSFRVenCode As String
        Dim lsSFRVenName As String

        Dim lsSFVendGroup As String
        Dim lsSFStat As String
        Dim lsEmpName As String
        Dim lsGNAddCode As String

        Dim lsSFStatus As String
        Dim lsSFTaxable As String
        Dim lsServiceTax As String
        Dim lsServiceTaxPercentage As String

        Dim lsSFPaid As String
        Dim lsxpuBillNo As String
        Dim lsDedupBillNo As String
        Dim lsSerTaxNo As String

        Dim lsPanNo As String
        Dim lsCSTNo As String
        Dim lsLSTNo As String
        Dim lsWCTNo As String

        Dim lsVATNo As String
        Dim lsSFExpMonth As String
        Dim lsProvisionFlag As String
        Dim lsVendorGID As String

        Dim lsSFApproverGID As String
        Dim lsECFDate As String

        Dim lobjErrorDatatable As New DataTable


        With lobjErrorDatatable
            .Columns.Add("Rowid")
            .Columns.Add("br_code")
            .Columns.Add("AUTH_DT")
            .Columns.Add("SF_BATCHNO")
            .Columns.Add("SF_DOCNUMB")
            .Columns.Add("SF_NUMBBIL")
            .Columns.Add("SF_BILNUMB")
            .Columns.Add("SF_VENCODE")
            .Columns.Add("SF_VENNAME")
            .Columns.Add("sf_actual_vendor")
            .Columns.Add("SF_BILDATE")
            .Columns.Add("SF_BILSTDT")
            .Columns.Add("SF_BILEDDT")
            .Columns.Add("SF_BAMOUNT")
            .Columns.Add("SF_BILDISC")
            .Columns.Add("SF_NETAMNT")
            .Columns.Add("SF_RVENCODE")
            .Columns.Add("SF_RVENNAME")
            .Columns.Add("SF_VENDGRP")
            .Columns.Add("SF_STAT")
            .Columns.Add("EMPNAME")
            .Columns.Add("GN_ADDCODE")
            .Columns.Add("sf_Status")
            .Columns.Add("sf_staxable")
            .Columns.Add("sf_servtax")
            .Columns.Add("sf_staxper")
            .Columns.Add("sf_paid")
            .Columns.Add("xpuBillNo")
            .Columns.Add("dedupBillNo")
            .Columns.Add("sertaxno")
            .Columns.Add("pan_no")
            .Columns.Add("CSTno")
            .Columns.Add("LSTno")
            .Columns.Add("WCTno")
            .Columns.Add("VATno")
            .Columns.Add("sf_expmonth")
            .Columns.Add("provision_flag")
            .Columns.Add("vendorbranch_gid")
            .Columns.Add("SF_APPROVER_ID")

            .Columns.Add("Error")
            .Columns.Add("Error Note")
        End With

        For i As Integer = 0 To lobjDataTable.Rows.Count - 1
            lsErrMessage = ""
            lsErrNote = ""

            With lobjDataTable.Rows(i)
                lsRowID = QuoteFilter(.Item("Rowid").ToString)
                lsBrCode = QuoteFilter(.Item("br_code").ToString)
                lsAuthDt = QuoteFilter(.Item("AUTH_DT").ToString)
                lsSFBatchNo = QuoteFilter(.Item("SF_BATCHNO").ToString)
                lsSFDocNumb = QuoteFilter(.Item("SF_DOCNUMB").ToString)
                lsSFNumbil = QuoteFilter(.Item("SF_NUMBBIL").ToString)
                lsSFBillNumb = QuoteFilter(.Item("SF_BILNUMB").ToString)
                lsSFVendCode = QuoteFilter(.Item("SF_VENCODE").ToString)
                lsSFVenName = QuoteFilter(.Item("SF_VENNAME").ToString)
                lsSFActualVendor = QuoteFilter(.Item("sf_actual_vendor").ToString)
                lsSFBillDate = QuoteFilter(.Item("SF_BILDATE").ToString)
                lsSFBillStDt = QuoteFilter(.Item("SF_BILSTDT").ToString)
                lsSFBillEDDt = QuoteFilter(.Item("SF_BILEDDT").ToString)
                lsSFBillAmount = QuoteFilter(.Item("SF_BAMOUNT").ToString)
                lsSFBillDesc = QuoteFilter(.Item("SF_BILDISC").ToString)
                lsSFNetAmount = QuoteFilter(.Item("SF_NETAMNT").ToString)
                lsSFRVenCode = QuoteFilter(.Item("SF_RVENCODE").ToString)
                lsSFRVenName = QuoteFilter(.Item("SF_RVENNAME").ToString)
                lsSFVendGroup = QuoteFilter(.Item("SF_VENDGRP").ToString)
                lsSFStat = QuoteFilter(.Item("SF_STAT").ToString)
                lsEmpName = QuoteFilter(.Item("EMPNAME").ToString)
                lsGNAddCode = QuoteFilter(.Item("GN_ADDCODE").ToString)
                lsSFStatus = QuoteFilter(.Item("sf_Status").ToString)
                lsSFTaxable = QuoteFilter(.Item("sf_staxable").ToString)
                lsServiceTax = QuoteFilter(.Item("sf_servtax").ToString)
                lsServiceTaxPercentage = QuoteFilter(.Item("sf_staxper").ToString)
                lsSFPaid = QuoteFilter(.Item("sf_paid").ToString)
                lsxpuBillNo = QuoteFilter(.Item("xpuBillNo").ToString)
                lsDedupBillNo = QuoteFilter(.Item("dedupBillNo").ToString)
                lsSerTaxNo = QuoteFilter(.Item("sertaxno").ToString)
                lsPanNo = QuoteFilter(.Item("pan_no").ToString)
                lsCSTNo = QuoteFilter(.Item("CSTno").ToString)
                lsLSTNo = QuoteFilter(.Item("LSTno").ToString)
                lsWCTNo = QuoteFilter(.Item("WCTno").ToString)
                lsVATNo = QuoteFilter(.Item("VATno").ToString)
                lsSFExpMonth = QuoteFilter(.Item("sf_expmonth").ToString)
                lsProvisionFlag = QuoteFilter(.Item("provision_flag").ToString)
                lsVendorGID = QuoteFilter(.Item("vendorbranch_gid").ToString)
                lsSFApproverGID = QuoteFilter(.Item("SF_APPROVER_ID").ToString)

            End With

            If lsSFBillDate <> "" Then
                If IsDate(lsSFBillDate) Then
                    lsSFBillDate = "'" & Format(CDate(lsSFBillDate), "yyyy-MM-dd") & "'"
                Else
                    lsSFBillDate = "NULL"
                End If
            Else
                lsSFBillDate = "NULL"
            End If

            If lsSFExpMonth <> "" Then
                If IsDate(lsSFExpMonth) Then
                    lsSFExpMonth = "'" & Format(CDate(lsSFExpMonth), "yyyy-MM-dd") & "'"
                Else
                    lsSFExpMonth = "NULL"
                End If
            Else
                lsSFExpMonth = "NULL"
            End If

            If lsSFDocNumb <> "" Then _
                lsECFDate = "20" & Mid(lsSFDocNumb, 4, 2) & "-" & Mid(lsSFDocNumb, 6, 2) & "-" & Mid(lsSFDocNumb, 8, 2)


            Dim lsSupplierEmployee As String
            Dim liPayToGID As Integer
            Dim liRaiserGID As Integer
            Dim liECFGID As Integer

            ' EMPLOYEE GID 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT employee_gid FROM iem_mst_temployee "
            lsSQL &= " WHERE employee_code='" & lsSFVendCode & "' "

            liRaiserGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            lsSupplierEmployee = IIf(lsSFVenName.Trim.ToLower = lsSFActualVendor.Trim.ToLower, "E", "S")

            If lsSupplierEmployee = "E" Then
                liPayToGID = liRaiserGID
            Else
                ' SUPPLIER GID 
                '-------------------------------------------------------------------------------------------------------------------------------
                lsSQL = ""
                lsSQL = " SELECT supplierheader_gid FROM asms_trn_tsupplierheader "
                lsSQL &= " WHERE Lower(supplierheader_name)='" & lsSFActualVendor.Trim.ToLower & "' "

                liPayToGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)
            End If

            Select Case 0
                Case liRaiserGID
                    lsErrNote &= " RAISER NOT FOUND : "
                Case liPayToGID
                    lsErrNote &= " PAY TO: SUPPLIER / VENDOR NOT FOUND: "
            End Select

            If lsErrNote <> "" Then GoTo NextFetch

            ' ECF HEADER GID FINDING 
            '-------------------------------------------------------------------------------------------------------------------------------
            lsSQL = ""
            lsSQL = " SELECT ecf_gid FROM iem_trn_tecf "
            lsSQL &= " WHERE ecf_no='" & lsSFDocNumb & "' "

            liECFGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


            ''            If liECFGID = 0 Then

            ''                ' INSERTING DATA INTO ECF HEADER 
            ''                '-------------------------------------------------------------------------------------------------------------------------------
            ''                lsSQL = ""
            ''                lsSQL &= " INSERT INTO iem_trn_tecf (ecf_supplier_employee, ecf_supplier_gid,ecf_employee_gid,ecf_date,ecf_no,ecf_amount"
            ''                lsSQL &= " ,ecf_isremoved,ecf_insert_by,ecf_insert_date,ecf_create_mode,ecf_raiser,ecf_doctype_gid,ecf_docsubtype_gid,ecf_claim_month"
            ''                lsSQL &= " ,ecf_currency_code,ecf_currency_rate,ecf_currency_amount,ecf_delmat_amount) values("
            ''"               lsSQL &= "'" & lsSupplierEmployee & "',


            ''                lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)

            ''                lsSQL = ""
            ''                lsSQL = " SELECT ecf_gid FROM iem_trn_tecf "
            ''                lsSQL &= " WHERE ecf_no='" & lsSFDocNumb & "' "

            ''                liECFGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            ''            End If

            ''            If liCBFGID <> 0 Then

            ''                lsSQL = ""
            ''                lsSQL = " SELECT cbfdetails_gid FROM fb_trn_tcbfdetails "
            ''                lsSQL &= " WHERE cbfdetails_cbfhead_gid='" & liCBFGID & "' AND cbfdetails_sno = " & lsDetailSno

            ''                liCBFDetailsGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            ''                If liCBFDetailsGID <> 0 Then
            ''                    lsSQL = ""
            ''                    lsSQL &= " DELETE FROM fb_trn_tcbfdetails "
            ''                    lsSQL &= " WHERE cbfdetails_cbfhead_gid='" & liCBFGID & "' AND cbfdetails_sno = " & lsDetailSno

            ''                    loDBConnection.ExecuteNonQuerySQL(lsSQL)
            ''                End If

            ''                ' INSERTING DATA INTO CBF DETAILS 
            ''                '-------------------------------------------------------------------------------------------------------------------------------
            ''                lsSQL = ""
            ''                lsSQL &= " INSERT INTO fb_trn_tcbfdetails(cbfheader_cbfobf_flag,cbfdetails_cbfhead_gid,cbfdetails_parprdesc,cbfdetails_year,cbfdetails_prod_gid,"
            ''                lsSQL &= " cbfdetails_desc,cbfdetails_uom_gid,cbfdetails_qty,cbfdetails_unitprice,cbfdetails_totalamt,cbfdetails_remarks,cbfdetails_chartofacc,cbfdetails_fccc,"
            ''                lsSQL &= " cbfdetails_budgetline,cbfdetails_budgetowner_gid,cbfdetails_vendor_gid,cbfdetails_prpardel_gid,cbfdetails_prodservgrp_gid,"
            ''                lsSQL &= " cbfdetails_sno) "

            ''                lsSQL &= " VALUES ('" & Mid(lsCBFOBFFlag.Trim, 1, 1) & "'," & liCBFGID & ",'" & lsPARPRDescription & "',''," & liProductGID & ",'" & lsCBFDetailsDescription & "'," & liUOMGID & "," & Val(lsQty)
            ''                lsSQL &= "," & Val(lsUnitPrice) & "," & Val(lsTotalAmount) & ",'" & lsCBFDetailsRemarks & "','" & lsCOA & "','" & Mid(lsFCCC.Trim, 1, 5) & "'," & lsBudgetLine & "," & liBudgetOwnerGId & "," & liVendorGID & ",0," & liProdServiceGID
            ''                lsSQL &= "," & lsDetailSno & ")"

            ''                lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)
            ''            Else
            ''                lsErrNote &= " CBF NOT FOUND "
            ''            End If

NextFetch:
            If lsErrMessage = "" And lsErrNote = "" Then
                liValid += 1
            Else
                liError += 1

                lobjErrorDatatable.Rows.Add()

                With lobjErrorDatatable.Rows(lobjErrorDatatable.Rows.Count - 1)
                    .Item("Rowid") = lsRowID
                    .Item("br_code") = lsBrCode
                    .Item("AUTH_DT") = lsAuthDt
                    .Item("SF_BATCHNO") = lsSFBatchNo
                    .Item("SF_DOCNUMB") = lsSFDocNumb
                    .Item("SF_NUMBBIL") = lsSFNumbil
                    .Item("SF_BILNUMB") = lsSFBillNumb
                    .Item("SF_VENCODE") = lsSFVendCode
                    .Item("SF_VENNAME") = lsSFVenName
                    .Item("sf_actual_vendor") = lsSFActualVendor
                    .Item("SF_BILDATE") = lsSFBillDate
                    .Item("SF_BILSTDT") = lsSFBillStDt
                    .Item("SF_BILEDDT") = lsSFBillEDDt
                    .Item("SF_BAMOUNT") = lsSFBillAmount
                    .Item("SF_BILDISC") = lsSFBillDesc
                    .Item("SF_NETAMNT") = lsSFNetAmount
                    .Item("SF_RVENCODE") = lsSFRVenCode
                    .Item("SF_RVENNAME") = lsSFRVenName
                    .Item("SF_VENDGRP") = lsSFVendGroup
                    .Item("SF_STAT") = lsSFStat
                    .Item("EMPNAME") = lsEmpName
                    .Item("GN_ADDCODE") = lsGNAddCode
                    .Item("sf_Status") = lsSFStatus
                    .Item("sf_staxable") = lsSFTaxable
                    .Item("sf_servtax") = lsServiceTax
                    .Item("sf_staxper") = lsServiceTaxPercentage
                    .Item("sf_paid") = lsSFPaid
                    .Item("xpuBillNo") = lsxpuBillNo
                    .Item("dedupBillNo") = lsDedupBillNo
                    .Item("sertaxno") = lsSerTaxNo
                    .Item("pan_no") = lsPanNo
                    .Item("CSTno") = lsCSTNo
                    .Item("LSTno") = lsLSTNo
                    .Item("WCTno") = lsWCTNo
                    .Item("VATno") = lsVATNo
                    .Item("sf_expmonth") = lsSFExpMonth
                    .Item("provision_flag") = lsProvisionFlag
                    .Item("vendorbranch_gid") = lsVendorGID
                    .Item("SF_APPROVER_ID") = lsSFApproverGID

                    .Item("Error") = lsErrMessage
                    .Item("Error Note") = lsErrNote
                End With
            End If

            txtStatus.Text = liValid & " Of " & lobjDataTable.Rows.Count & " Records Migrated...  Error - " & liError
            Application.DoEvents()
        Next

        If lobjErrorDatatable.Rows.Count > 0 Then
            PrintDGridviewXML(lobjErrorDatatable, Application.StartupPath & "\CBF-MIG-ERR.xls")
            MessageBox.Show("Descrepance Records Spooled @ " & Application.StartupPath & "\CBF-MIG-ERR.xls")
        Else
            MessageBox.Show("Imported successfully...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If


        'Catch ex As Exception
        '    grpMain.Enabled = True
        '    MsgBox(ex.Message, MsgBoxStyle.Critical)
        'End Try
    End Sub

    '' ''    Public Sub CBF_Data_Migration()
    '' ''        'Try
    '' ''        Dim lsSQL As String = ""
    '' ''        Dim lsFileName As String = ""
    '' ''        Dim lsResult As String = ""
    '' ''        Dim liDumpGid As Integer = 0
    '' ''        Dim lobjOleDbConnection As New OleDb.OleDbConnection
    '' ''        Dim lobjDataAdapter As OleDb.OleDbDataAdapter
    '' ''        Dim lobjDataTable As DataTable
    '' ''        Dim lobjDataSet As New DataSet
    '' ''        Dim liError As Integer
    '' ''        Dim liValid As Integer

    '' ''        Dim lsErrMessage As String = ""
    '' ''        Dim lsErrNote As String = ""


    '' ''        lsFileName = Mid(txtFileName.Text.Trim, InStrRev(txtFileName.Text.Trim, "\") + 1)

    '' ''        Try
    '' ''            If lobjOleDbConnection.State = 1 Then lobjOleDbConnection.Close()
    '' ''            With lobjOleDbConnection

    '' ''                If Microsoft.VisualBasic.Right(lsFileName, 4) = "xlsx" Then
    '' ''                    'read a 2007 file   
    '' ''                    .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
    '' ''                                         txtFileName.Text.Trim & ";" + ";Extended Properties='Excel 12.0 Xml;HDR=YES';"
    '' ''                Else
    '' ''                    'read a 97-2003 file   
    '' ''                    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" & _
    '' ''                                            txtFileName.Text.Trim & ";" + "Extended Properties=Excel 8.0;"
    '' ''                End If

    '' ''                .Open()

    '' ''            End With

    '' ''        Catch ex As Exception
    '' ''            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '' ''            lobjOleDbConnection.Close()
    '' ''            Exit Sub
    '' ''        End Try

    '' ''        lsSQL = "SELECT * FROM [" & cmbSheet.Text & "$]"

    '' ''        lobjDataAdapter = New OleDb.OleDbDataAdapter(lsSQL, lobjOleDbConnection)
    '' ''        lobjDataAdapter.Fill(lobjDataSet, "DATA")
    '' ''        lobjDataTable = lobjDataSet.Tables("DATA")
    '' ''        lobjOleDbConnection.Close()

    '' ''        Dim liIndex As Integer

    '' ''        For liIndex = 0 To lobjDataTable.Columns.Count - 1
    '' ''            If Not InStr(fsColumnHeaders.ToLower.Trim, "," & lobjDataTable.Columns(liIndex).ColumnName.ToString.Trim.ToLower & ",") > 0 Then
    '' ''                MessageBox.Show("ECF Dump is not in the correct format...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
    '' ''                Exit Sub
    '' ''            End If
    '' ''        Next

    '' ''        Dim lsCBFNumber As String
    '' ''        Dim lsDetailSno As String
    '' ''        Dim lsCBFOBFFlag As String
    '' ''        Dim lsStartDate As String
    '' ''        Dim lsEndDate As String
    '' ''        Dim lsProjectOwner As String
    '' ''        Dim lsBranchCode As String
    '' ''        Dim lsMode As String
    '' ''        Dim lsApprovalType As String
    '' ''        Dim lsIsBudgeted As String
    '' ''        Dim lsDeviationAmount As String
    '' ''        Dim lsCBFAmount As String
    '' ''        Dim lsDescription As String
    '' ''        Dim lsRaiser As String
    '' ''        Dim lsRequestFor As String
    '' ''        Dim lsRemarks As String
    '' ''        Dim lsIsBranchSingle As String
    '' ''        Dim lsBudgetOwner As String
    '' ''        Dim lsPARPRDescription As String
    '' ''        Dim lsProductService As String
    '' ''        Dim lsCBFDetailsDescription As String
    '' ''        Dim lsUOM As String
    '' ''        Dim lsQty As String
    '' ''        Dim lsUnitPrice As String
    '' ''        Dim lsTotalAmount As String
    '' ''        Dim lsCBFDetailsRemarks As String
    '' ''        Dim lsCOA As String
    '' ''        Dim lsFCCC As String
    '' ''        Dim lsBudgetLine As String
    '' ''        Dim lsVendor As String
    '' ''        Dim lsProductGroup As String

    '' ''        Dim liCBFGID As Integer
    '' ''        Dim liProductGID As Integer
    '' ''        Dim liUOMGID As Integer
    '' ''        Dim liVendorGID As Integer
    '' ''        Dim liProdServiceGID As Integer
    '' ''        Dim liCBFDetailsGID As Integer

    '' ''        Dim liRaisorGID As Integer
    '' ''        Dim liProjectOwnerGID As Integer
    '' ''        Dim liBudgetOwnerGId As Integer

    '' ''        Dim liBranchGID As Integer

    '' ''        Dim lobjErrorDatatable As New DataTable


    '' ''        With lobjErrorDatatable
    '' ''            .Columns.Add("CBFNo")
    '' ''            .Columns.Add("Detail_Sno")
    '' ''            .Columns.Add("CBF_OBF_Flag")
    '' ''            .Columns.Add("Start_Date")
    '' ''            .Columns.Add("End_Date")
    '' ''            .Columns.Add("Project_Owner")
    '' ''            .Columns.Add("Branch")
    '' ''            .Columns.Add("Mode")
    '' ''            .Columns.Add("Approval_Type")
    '' ''            .Columns.Add("Is_Budgeted")
    '' ''            .Columns.Add("Deviation_Amount")
    '' ''            .Columns.Add("CBF_Amount")

    '' ''            .Columns.Add("Description")
    '' ''            .Columns.Add("Raiser")
    '' ''            .Columns.Add("Request_For")
    '' ''            .Columns.Add("Remarks")
    '' ''            .Columns.Add("Is_Branch_Single")
    '' ''            .Columns.Add("Budget_Owner")
    '' ''            .Columns.Add("PAR_PR_Description")
    '' ''            .Columns.Add("Product_Service")
    '' ''            .Columns.Add("CBF_Details_Description")
    '' ''            .Columns.Add("UOM")
    '' ''            .Columns.Add("QTY")
    '' ''            .Columns.Add("Unit_Price")

    '' ''            .Columns.Add("Total_Amount")
    '' ''            .Columns.Add("CBF_Details_Remarks")
    '' ''            .Columns.Add("COA")
    '' ''            .Columns.Add("FCCC")
    '' ''            .Columns.Add("Budget_Line")
    '' ''            .Columns.Add("Vendor")
    '' ''            .Columns.Add("Product_Group")

    '' ''            .Columns.Add("Error")
    '' ''            .Columns.Add("Error Note")
    '' ''        End With

    '' ''        For i As Integer = 0 To lobjDataTable.Rows.Count - 1
    '' ''            lsErrMessage = ""
    '' ''            lsErrNote = ""

    '' ''            With lobjDataTable.Rows(i)

    '' ''                lsCBFNumber = QuoteFilter(.Item("CBFNo").ToString)
    '' ''                lsDetailSno = QuoteFilter(.Item("Detail_Sno").ToString)
    '' ''                lsCBFOBFFlag = QuoteFilter(.Item("CBF_OBF_Flag").ToString)
    '' ''                lsStartDate = QuoteFilter(.Item("Start_Date").ToString)
    '' ''                lsEndDate = QuoteFilter(.Item("End_Date").ToString)
    '' ''                lsProjectOwner = QuoteFilter(.Item("Project_Owner").ToString)
    '' ''                lsBranchCode = QuoteFilter(.Item("Branch").ToString)
    '' ''                lsMode = QuoteFilter(.Item("Mode").ToString)
    '' ''                lsApprovalType = QuoteFilter(.Item("Approval_Type").ToString)
    '' ''                lsIsBudgeted = QuoteFilter(.Item("Is_Budgeted").ToString)
    '' ''                lsDeviationAmount = QuoteFilter(.Item("Deviation_Amount").ToString)
    '' ''                lsCBFAmount = QuoteFilter(.Item("CBF_Amount").ToString)

    '' ''                lsDescription = QuoteFilter(.Item("Description").ToString)
    '' ''                lsRaiser = QuoteFilter(.Item("Raiser").ToString)
    '' ''                lsRequestFor = QuoteFilter(.Item("Request_For").ToString)
    '' ''                lsRemarks = QuoteFilter(.Item("Remarks").ToString)
    '' ''                lsIsBranchSingle = QuoteFilter(.Item("Is_Branch_Single").ToString)
    '' ''                lsBudgetOwner = QuoteFilter(.Item("Budget_Owner").ToString)
    '' ''                lsPARPRDescription = QuoteFilter(.Item("PAR_PR_Description").ToString)
    '' ''                lsProductService = QuoteFilter(.Item("Product_Service").ToString)
    '' ''                lsCBFDetailsDescription = QuoteFilter(.Item("CBF_Details_Description").ToString)
    '' ''                lsUOM = QuoteFilter(.Item("UOM").ToString)
    '' ''                lsQty = QuoteFilter(.Item("QTY").ToString)
    '' ''                lsUnitPrice = QuoteFilter(.Item("Unit_Price").ToString)

    '' ''                lsTotalAmount = QuoteFilter(.Item("Total_Amount").ToString)
    '' ''                lsCBFDetailsRemarks = QuoteFilter(.Item("CBF_Details_Remarks").ToString)
    '' ''                lsCOA = QuoteFilter(.Item("Chart_Of_Acc").ToString)
    '' ''                lsFCCC = QuoteFilter(.Item("FCCC").ToString)
    '' ''                lsBudgetLine = QuoteFilter(.Item("Budget_Line").ToString)
    '' ''                lsVendor = QuoteFilter(.Item("Vendor").ToString)
    '' ''                lsProductGroup = QuoteFilter(.Item("Product_Group").ToString)
    '' ''            End With

    '' ''            If lsStartDate <> "" Then
    '' ''                If IsDate(lsStartDate) Then
    '' ''                    lsStartDate = "'" & Format(CDate(lsStartDate), "yyyy-MM-dd") & "'"
    '' ''                Else
    '' ''                    lsStartDate = "NULL"
    '' ''                End If
    '' ''            Else
    '' ''                lsStartDate = "NULL"
    '' ''            End If

    '' ''            If lsEndDate <> "" Then
    '' ''                If IsDate(lsEndDate) Then
    '' ''                    lsEndDate = "'" & Format(CDate(lsEndDate), "yyyy-MM-dd") & "'"
    '' ''                Else
    '' ''                    lsEndDate = "NULL"
    '' ''                End If
    '' ''            Else
    '' ''                lsEndDate = "NULL"
    '' ''            End If


    '' ''            liCBFGID = 0
    '' ''            liProductGID = 0
    '' ''            liUOMGID = 0
    '' ''            liVendorGID = 0
    '' ''            liProdServiceGID = 0
    '' ''            liCBFDetailsGID = 0
    '' ''            liRaisorGID = 0
    '' ''            liProjectOwnerGID = 0
    '' ''            liBudgetOwnerGId = 0
    '' ''            liBranchGID = 0


    '' ''            ' UOM GID 
    '' ''            '-------------------------------------------------------------------------------------------------------------------------------
    '' ''            lsSQL = ""
    '' ''            lsSQL = " SELECT uom_gid FROM iem_mst_tuom "
    '' ''            lsSQL &= " WHERE LOWER(uom_code)='" & lsUOM.ToLower.Trim & "'"

    '' ''            liUOMGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


    '' ''            ' PRODUCT / SERVICE GID 
    '' ''            '-------------------------------------------------------------------------------------------------------------------------------
    '' ''            If InStr(lsProductService, "-") = 0 Then _
    '' ''                lsProductService &= " - "

    '' ''            lsSQL = ""
    '' ''            lsSQL = " SELECT prodservice_gid FROM fb_mst_tprodservice "
    '' ''            lsSQL &= " WHERE prodservice_code='" & Mid(lsProductService, 1, InStr(lsProductService, "-") - 1).Trim & "' "

    '' ''            liProductGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

    '' ''            lsSQL = ""
    '' ''            lsSQL = " SELECT prodservice_prodservicegid FROM fb_mst_tprodservice "
    '' ''            lsSQL &= " WHERE prodservice_code='" & Mid(lsProductService, 1, InStr(lsProductService, "-") - 1).Trim & "' "

    '' ''            liProdServiceGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


    '' ''            ' SUPPLIER GID 
    '' ''            '-------------------------------------------------------------------------------------------------------------------------------
    '' ''            lsSQL = ""
    '' ''            lsSQL = " SELECT supplierheader_gid FROM asms_trn_tsupplierheader "
    '' ''            lsSQL &= " WHERE supplierheader_name='" & lsVendor & "' "

    '' ''            liVendorGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


    '' ''            ' EMPLOYEE GID 
    '' ''            '-------------------------------------------------------------------------------------------------------------------------------
    '' ''            lsSQL = ""
    '' ''            lsSQL = " SELECT employee_gid FROM iem_mst_temployee "
    '' ''            lsSQL &= " WHERE employee_code='" & Mid(lsRaiser, 1, InStr(lsRaiser, "-") - 1).Trim & "' "

    '' ''            liRaisorGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

    '' ''            ' PROJECT OWNER GID 
    '' ''            '-------------------------------------------------------------------------------------------------------------------------------
    '' ''            lsSQL = ""
    '' ''            lsSQL = " SELECT employee_gid FROM iem_mst_temployee "
    '' ''            lsSQL &= " WHERE employee_code='" & Mid(lsProjectOwner, 1, InStr(lsProjectOwner, "-") - 1).Trim & "' "

    '' ''            liProjectOwnerGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

    '' ''            ' BUDGET OWNER GID 
    '' ''            '-------------------------------------------------------------------------------------------------------------------------------
    '' ''            lsSQL = ""
    '' ''            lsSQL = " SELECT employee_gid FROM iem_mst_temployee "
    '' ''            lsSQL &= " WHERE employee_code='" & Mid(lsBudgetOwner, 1, InStr(lsBudgetOwner, "-") - 1).Trim & "' "

    '' ''            liBudgetOwnerGId = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


    '' ''            ' BRANCH GID 
    '' ''            '-------------------------------------------------------------------------------------------------------------------------------
    '' ''            lsSQL = ""
    '' ''            lsSQL = " SELECT branch_gid FROM iem_mst_tbranch "
    '' ''            lsSQL &= " WHERE branch_code='" & Mid(lsBranchCode, 1, InStr(lsBranchCode, "-") - 1).Trim & "' "

    '' ''            liBranchGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


    '' ''            Select Case 0
    '' ''                Case liProductGID
    '' ''                    lsErrNote &= " PRODUCT NOT FOUND : "
    '' ''                Case liProdServiceGID
    '' ''                    lsErrNote &= " PRODUCT SERVICE NOT FOUND : "
    '' ''                Case liVendorGID
    '' ''                    lsErrNote &= " VENDOR NOT FOUND : "
    '' ''                Case liRaisorGID
    '' ''                    lsErrNote &= " RAISER NOT FOUND : "
    '' ''                Case liProjectOwnerGID
    '' ''                    lsErrNote &= " PROJECT OWNER NOT FOUND : "
    '' ''                Case liBudgetOwnerGId
    '' ''                    lsErrNote &= " BUDGET OWNER NOT FOUND : "
    '' ''                Case liBranchGID
    '' ''                    lsErrNote &= " BRANCH NOT FOUND : "
    '' ''            End Select

    '' ''            If lsErrNote <> "" Then GoTo NextFetch

    '' ''            ' CBF HEADER GID FINDING 
    '' ''            '-------------------------------------------------------------------------------------------------------------------------------
    '' ''            lsSQL = ""
    '' ''            lsSQL = " SELECT cbfheader_gid FROM fb_trn_tcbfheader "
    '' ''            lsSQL &= " WHERE cbfheader_cbfno='" & lsCBFNumber & "' "

    '' ''            liCBFGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)


    '' ''            If liCBFGID = 0 Then

    '' ''                ' INSERTING DATA INTO CBF HEADER 
    '' ''                '-------------------------------------------------------------------------------------------------------------------------------
    '' ''                lsSQL = ""
    '' ''                lsSQL &= " INSERT INTO fb_trn_tcbfheader(cbfheader_cbfno,cbfheader_cbfobf_flag,cbfheader_date,cbfheader_enddate,cbfheader_projectowner,cbfheader_branch_gid,"
    '' ''                lsSQL &= " cbfheader_mode,cbfheader_prpar_gid,cbfheader_approvaltype,cbfheader_isbudgeted,cbfheader_Devi_amt,cbfheader_cbfamt,cbfheader_desc,"
    '' ''                lsSQL &= " cbfheader_rasier_gid,cbfheader_requestfor_gid,cbfheader_requesttype,cbfheader_remarks,cbfheader_budgetowner_gid,cbfheader_status)  "

    '' ''                lsSQL &= " VALUES('" & lsCBFNumber & "','" & Mid(lsCBFOBFFlag.Trim, 1, 1) & "'," & lsStartDate & "," & lsEndDate & "," & liProjectOwnerGID & ","
    '' ''                lsSQL &= liBranchGID & ",'" & lsMode & "',0,'" & Mid(lsApprovalType, 1, 1) & "','" & lsIsBudgeted & "'," & lsDeviationAmount & "," & lsCBFAmount & ","
    '' ''                lsSQL &= "'" & lsDescription & "'," & liRaisorGID & "," & IIf(lsRequestFor.Trim = "IT", 3, 4) & ",'','" & lsRemarks.Trim & "'," & liBudgetOwnerGId & ",5) "

    '' ''                lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)

    '' ''                lsSQL = ""
    '' ''                lsSQL = " SELECT cbfheader_gid FROM fb_trn_tcbfheader "
    '' ''                lsSQL &= " WHERE cbfheader_cbfno='" & lsCBFNumber & "' "

    '' ''                liCBFGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

    '' ''            End If

    '' ''            If liCBFGID <> 0 Then

    '' ''                lsSQL = ""
    '' ''                lsSQL = " SELECT cbfdetails_gid FROM fb_trn_tcbfdetails "
    '' ''                lsSQL &= " WHERE cbfdetails_cbfhead_gid='" & liCBFGID & "' AND cbfdetails_sno = " & lsDetailSno

    '' ''                liCBFDetailsGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

    '' ''                If liCBFDetailsGID <> 0 Then
    '' ''                    lsSQL = ""
    '' ''                    lsSQL &= " DELETE FROM fb_trn_tcbfdetails "
    '' ''                    lsSQL &= " WHERE cbfdetails_cbfhead_gid='" & liCBFGID & "' AND cbfdetails_sno = " & lsDetailSno

    '' ''                    loDBConnection.ExecuteNonQuerySQL(lsSQL)
    '' ''                End If

    '' ''                ' INSERTING DATA INTO CBF DETAILS 
    '' ''                '-------------------------------------------------------------------------------------------------------------------------------
    '' ''                lsSQL = ""
    '' ''                lsSQL &= " INSERT INTO fb_trn_tcbfdetails(cbfheader_cbfobf_flag,cbfdetails_cbfhead_gid,cbfdetails_parprdesc,cbfdetails_year,cbfdetails_prod_gid,"
    '' ''                lsSQL &= " cbfdetails_desc,cbfdetails_uom_gid,cbfdetails_qty,cbfdetails_unitprice,cbfdetails_totalamt,cbfdetails_remarks,cbfdetails_chartofacc,cbfdetails_fccc,"
    '' ''                lsSQL &= " cbfdetails_budgetline,cbfdetails_budgetowner_gid,cbfdetails_vendor_gid,cbfdetails_prpardel_gid,cbfdetails_prodservgrp_gid,"
    '' ''                lsSQL &= " cbfdetails_sno) "

    '' ''                lsSQL &= " VALUES ('" & Mid(lsCBFOBFFlag.Trim, 1, 1) & "'," & liCBFGID & ",'" & lsPARPRDescription & "',''," & liProductGID & ",'" & lsCBFDetailsDescription & "'," & liUOMGID & "," & Val(lsQty)
    '' ''                lsSQL &= "," & Val(lsUnitPrice) & "," & Val(lsTotalAmount) & ",'" & lsCBFDetailsRemarks & "','" & lsCOA & "','" & Mid(lsFCCC.Trim, 1, 5) & "'," & lsBudgetLine & "," & liBudgetOwnerGId & "," & liVendorGID & ",0," & liProdServiceGID
    '' ''                lsSQL &= "," & lsDetailSno & ")"

    '' ''                lsErrMessage &= loDBConnection.ExecuteNonQuerySQL(lsSQL)
    '' ''            Else
    '' ''                lsErrNote &= " CBF NOT FOUND "
    '' ''            End If

    '' ''NextFetch:
    '' ''            If lsErrMessage = "" And lsErrNote = "" Then
    '' ''                liValid += 1
    '' ''            Else
    '' ''                liError += 1

    '' ''                lobjErrorDatatable.Rows.Add()

    '' ''                With lobjErrorDatatable.Rows(lobjErrorDatatable.Rows.Count - 1)

    '' ''                    .Item("CBFNo") = lsCBFNumber
    '' ''                    .Item("Detail_Sno") = lsDetailSno
    '' ''                    .Item("CBF_OBF_Flag") = lsCBFOBFFlag
    '' ''                    .Item("Start_Date") = lsStartDate
    '' ''                    .Item("End_Date") = lsEndDate
    '' ''                    .Item("Project_Owner") = lsProjectOwner
    '' ''                    .Item("Branch") = lsBranchCode
    '' ''                    .Item("Mode") = lsMode
    '' ''                    .Item("Approval_Type") = lsApprovalType
    '' ''                    .Item("Is_Budgeted") = lsIsBudgeted
    '' ''                    .Item("Deviation_Amount") = lsDeviationAmount
    '' ''                    .Item("CBF_Amount") = lsCBFAmount

    '' ''                    .Item("Description") = lsDescription
    '' ''                    .Item("Raiser") = lsRaiser
    '' ''                    .Item("Request_For") = lsRequestFor
    '' ''                    .Item("Remarks") = lsRemarks
    '' ''                    .Item("Is_Branch_Single") = lsIsBranchSingle
    '' ''                    .Item("Budget_Owner") = lsBudgetOwner
    '' ''                    .Item("PAR_PR_Description") = lsPARPRDescription
    '' ''                    .Item("Product_Service") = lsProductService
    '' ''                    .Item("CBF_Details_Description") = lsCBFDetailsDescription
    '' ''                    .Item("UOM") = lsUOM
    '' ''                    .Item("QTY") = lsQty
    '' ''                    .Item("Unit_Price") = lsUnitPrice

    '' ''                    .Item("Total_Amount") = lsTotalAmount
    '' ''                    .Item("CBF_Details_Remarks") = lsCBFDetailsRemarks
    '' ''                    .Item("COA") = lsCOA
    '' ''                    .Item("FCCC") = lsFCCC
    '' ''                    .Item("Budget_Line") = lsBudgetLine
    '' ''                    .Item("Vendor") = lsVendor
    '' ''                    .Item("Product_Group") = lsProductGroup

    '' ''                    .Item("Error") = lsErrMessage
    '' ''                    .Item("Error Note") = lsErrNote
    '' ''                End With
    '' ''            End If

    '' ''            txtStatus.Text = liValid & " Of " & lobjDataTable.Rows.Count & " Records Migrated...  Error - " & liError
    '' ''            Application.DoEvents()
    '' ''        Next

    '' ''        If lobjErrorDatatable.Rows.Count > 0 Then
    '' ''            PrintDGridviewXML(lobjErrorDatatable, Application.StartupPath & "\CBF-MIG-ERR.xls")
    '' ''            MessageBox.Show("Descrepance Records Spooled @ " & Application.StartupPath & "\CBF-MIG-ERR.xls")
    '' ''        Else
    '' ''            MessageBox.Show("Imported successfully...", gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
    '' ''        End If


    '' ''        'Catch ex As Exception
    '' ''        '    grpMain.Enabled = True
    '' ''        '    MsgBox(ex.Message, MsgBoxStyle.Critical)
    '' ''        'End Try
    '' ''    End Sub

    Private Sub frmFeedbackImport_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Return Then
            SendKeys.Send("{TAB}")
        End If
    End Sub

    Private Sub frmFeedbackImport_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.CenterToScreen()
        'Me.CenterToParent()



        'loDBConnection.OpenConnection("WIN-M0GE46V6843", "sa", "gnsa", "iem_mig")
        loDBConnection.OpenConnection("WIN-M0GE46V6843", "sa", "gnsa", "iem_8540")

    End Sub

    Private Sub CreateTemplate(ByVal lsHeaders As String(), ByVal lsTemplateName As String)
        'Dim xlApp As New Excel.Application
        'Dim xlBook As Excel.Workbook
        'Dim xlSheet As New Excel.Worksheet
        'Dim iCol As Integer

        'Const cFirstRow = 1

        'Try
        '    xlBook = xlApp.Workbooks.Add
        '    xlSheet = xlApp.ActiveSheet
        '    xlSheet.Name = "Sheet1"
        '    xlApp.Visible = True
        '    With xlSheet
        '        For iCol = 1 To UBound(lsHeaders) + 1
        '            .Cells(cFirstRow, iCol) = lsHeaders(iCol - 1)
        '            .Cells(cFirstRow, iCol).Font.Bold = True
        '            .Columns(iCol).AutoFit()
        '        Next
        '        .Name = lsTemplateName
        '    End With
        '    xlBook.Sheets("Sheet2").Delete()
        '    xlBook.Sheets("Sheet3").Delete()
        'Catch ex As Exception
        '    MessageBox.Show(ex.Message, gsTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try
    End Sub

    Private Sub btnTemplate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTemplate.Click
        Dim lsHeaders, lsFormatHeaders() As String
        lsHeaders = Mid(fsColumnHeaders, 2, InStrRev(fsColumnHeaders, ",") - 2)
        lsFormatHeaders = lsHeaders.Split(",")
        Call CreateTemplate(lsFormatHeaders, "Template")
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
End Class