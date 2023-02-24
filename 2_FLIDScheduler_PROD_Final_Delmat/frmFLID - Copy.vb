Imports System.Data.SqlClient

Public Class frmFLID

    Dim loDBConnection As New iODBCconnection

    'DESIGNATION
    'EMPLOYEE
    'REGION

    Dim liRegionID As Integer
    Dim liCityID As Integer
    Dim liCountryID As Integer
    Dim liStateID As Integer

    Dim lsLinkedServer As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'loDBConnection.OpenConnection("GNSA_FLEXICODE", "sa", "asng", "ficc_iem_uat")
        loDBConnection.OpenConnection("KATHIR-PC\SQLEXPRESS", "sa", "gnsa", "ficc_iem")

        lsLinkedServer = "[192.168.84.81\INST_LMSUAT].[New_Adrelin_FinalMgmtDemo]."

        Me.Text = "BANK "
        Application.DoEvents()

        FLIDS_BANK()


        Me.Text = "FCCC"
        Application.DoEvents()

        FLIDS_FCCC()


        Me.Text = "OU"
        Application.DoEvents()

        FLIDS_OU()

        Me.Text = "PRODUCT "
        Application.DoEvents()

        FLIDS_PRODUCT()


        Me.Text = "REGION"
        Application.DoEvents()

        FLIDS_REGION()


        Me.Text = "COUNTRY"
        Application.DoEvents()

        FLIDS_COUNTRY()


        Me.Text = "GRADE"
        Application.DoEvents()

        FLIDS_GRADE()


        Me.Text = "DEPARTMENT"
        Application.DoEvents()

        FLIDS_DEPARTMENT()

        Me.Text = "DESIGNATION "
        Application.DoEvents()

        FLIDS_DESIGNATION()


        Me.Text = "HOLIDAY"
        Application.DoEvents()

        FLIDS_HOLIDAY()

        FLIDS_EMPLOYEE()

        MsgBox("Success")

    End Sub

    Private Sub FLIDS_EMPLOYEE()
        Dim loMigrationData As SqlDataReader

        Dim lsLastModifiedOn As String
        Dim lsTableName As String
        Dim lsSQL As String
        Dim liNewCount, liModifiedCount As Integer

        Dim EMP_STAFFID As String
        Dim EMP_COMPANYNAME As String
        Dim DATE_OF_BIRTH As String
        Dim PER_GENDER As String
        Dim CON_HOME_ADDRESS1 As String
        Dim CON_HOME_ADDRESS2 As String
        Dim CON_HOME_ADDRESS3 As String
        Dim CON_HOME_ADDRESS4 As String
        Dim CITY_NAME As String
        Dim STATE_NAME As String
        Dim REGION_NAME As String
        Dim COUNTRY_NAME As String
        Dim PIN_CODE As String
        Dim DATE_OF_JOINING As String
        Dim GRADE_NAME As String
        Dim HRIS_DESIGNATION As String
        Dim IEM_DESIGNATION As String
        Dim DEPARTMENT As String
        Dim BUSINESS As String
        Dim OFFICE_MAIL_ID As String
        Dim PERSONAL_EMAIL_ID As String
        Dim CONTACT_NO As String
        Dim MOBILE_NO As String
        Dim EMP_ACCOUNTNO As String
        Dim EMP_BANKNAME As String
        Dim EMP_IFS_NO As String
        Dim BRANCH_CODE As String
        Dim BSCC_NAME As String
        Dim EMP_REPORTINGTO As String
        Dim DATE_OF_RESIGNATION As String
        Dim EMPLOYEE_STATUS As String
        Dim PRODUCT As String
        Dim OU_NAME As String

        Dim liGradeGID As Integer
        Dim liDesignationGID As Integer
        Dim liDepartmentGID As Integer


        Dim liBusinessGID As Integer
        Dim liBankGID As Integer
        Dim liFCCCGID As Integer
        Dim liProductGID As Integer
        Dim liOUGID As Integer
        Dim liBranchGID As Integer

        Dim liErrored As Integer

        liErrored = 0
        Dim lbIsModified As Boolean


        lsSQL = ""
        lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
        lsSQL &= " FROM iem_mst_temployee "

        lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL)

        lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)

        lsTableName = " [econnect].[dbo].[IEM_EMPLOYEE_FIELDS_FIC] "
        liNewCount = 0
        liModifiedCount = 0

        lsSQL = ""
        lsSQL &= " SELECT * FROM " & lsTableName
        lsSQL &= " WHERE LAST_MODIFIED_DATE IS NULL "
        lsSQL &= " OR LAST_MODIFIED_DATE >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "

        loMigrationData = loDBConnection.GetDataReader(lsSQL)

        If loMigrationData.HasRows Then

            While loMigrationData.Read

                EMP_STAFFID = loMigrationData.Item("EMP_STAFFID").ToString.Trim

                EMP_COMPANYNAME = loMigrationData.Item("EMP_COMPANYNAME").ToString.Trim
                EMP_COMPANYNAME = FormatTextInput(EMP_COMPANYNAME)
                EMP_COMPANYNAME = Mid(EMP_COMPANYNAME, 1, 64)

                DATE_OF_BIRTH = loMigrationData.Item("DATE_OF_BIRTH")

                PER_GENDER = loMigrationData.Item("PER_GENDER").ToString.Trim
                PER_GENDER = Mid(PER_GENDER, 1, 1)

                CON_HOME_ADDRESS1 = loMigrationData.Item("CON_HOME_ADDRESS1").ToString.Trim
                CON_HOME_ADDRESS1 = FormatTextInput(CON_HOME_ADDRESS1)
                CON_HOME_ADDRESS1 = Mid(CON_HOME_ADDRESS1, 1, 64)

                CON_HOME_ADDRESS2 = loMigrationData.Item("CON_HOME_ADDRESS2").ToString.Trim
                CON_HOME_ADDRESS2 = FormatTextInput(CON_HOME_ADDRESS2)
                CON_HOME_ADDRESS2 = Mid(CON_HOME_ADDRESS2, 1, 64)


                CON_HOME_ADDRESS3 = loMigrationData.Item("CON_HOME_ADDRESS3").ToString.Trim
                CON_HOME_ADDRESS3 = FormatTextInput(CON_HOME_ADDRESS3)
                CON_HOME_ADDRESS3 = Mid(CON_HOME_ADDRESS3, 1, 64)

                CON_HOME_ADDRESS4 = loMigrationData.Item("CON_HOME_ADDRESS4").ToString.Trim
                CON_HOME_ADDRESS4 = FormatTextInput(CON_HOME_ADDRESS4)
                CON_HOME_ADDRESS4 = Mid(CON_HOME_ADDRESS4, 1, 64)

                liRegionID = 0
                liCountryID = 0
                liStateID = 0
                liCityID = 0

                REGION_NAME = loMigrationData.Item("REGION_NAME").ToString.Trim
                REGION_NAME = Mid(REGION_NAME, 1, 16)
                liRegionID = GetRegionGID(REGION_NAME)
                REGION_NAME = FormatTextInput(REGION_NAME)

                COUNTRY_NAME = loMigrationData.Item("COUNTRY_NAME").ToString.Trim
                COUNTRY_NAME = Mid(COUNTRY_NAME, 1, 64)
                liCountryID = GetCountryGID(COUNTRY_NAME)
                COUNTRY_NAME = FormatTextInput(COUNTRY_NAME)


                STATE_NAME = loMigrationData.Item("STATE_NAME").ToString.Trim
                STATE_NAME = Mid(STATE_NAME, 1, 64)
                liStateID = GetStateGID(STATE_NAME, liRegionID, liCountryID)
                STATE_NAME = FormatTextInput(STATE_NAME)

                CITY_NAME = loMigrationData.Item("CITY_NAME").ToString.Trim
                CITY_NAME = Mid(CITY_NAME, 1, 64)
                liCityID = GetCityGID(CITY_NAME, liRegionID, liCountryID, liStateID)
                CITY_NAME = FormatTextInput(CITY_NAME)


                PIN_CODE = loMigrationData.Item("PIN_CODE").ToString.Trim
                PIN_CODE = Mid(PIN_CODE, 1, 8)

                DATE_OF_JOINING = loMigrationData.Item("DATE_OF_JOINING")

                GRADE_NAME = loMigrationData.Item("GRADE_NAME").ToString.Trim
                GRADE_NAME = Mid(GRADE_NAME, 1, 32)
                liGradeGID = GetGradeGID(GRADE_NAME)

                lsSQL = ""
                lsSQL &= " SELECT grade_code "
                lsSQL &= " FROM iem_mst_tgrade WHERE grade_gid=" & liGradeGID

                GRADE_NAME = loDBConnection.GetExecuteScalar(lsSQL)



                HRIS_DESIGNATION = loMigrationData.Item("HRIS_DESIGNATION").ToString.Trim
                HRIS_DESIGNATION = Mid(HRIS_DESIGNATION, 1, 32)
                HRIS_DESIGNATION = FormatTextInput(HRIS_DESIGNATION)

                IEM_DESIGNATION = loMigrationData.Item("IEM_DESIGNATION").ToString.Trim
                IEM_DESIGNATION = Mid(IEM_DESIGNATION, 1, 32)
                liDesignationGID = GetDesignationGID(HRIS_DESIGNATION)
                IEM_DESIGNATION = FormatTextInput(IEM_DESIGNATION)

                DEPARTMENT = loMigrationData.Item("DEPARTMENT").ToString.Trim
                DEPARTMENT = Mid(DEPARTMENT, 1, 32)
                liDepartmentGID = GetDepartmentGID(DEPARTMENT)
                DEPARTMENT = FormatTextInput(DEPARTMENT)

                BUSINESS = loMigrationData.Item("BUSINESS").ToString.Trim
                BUSINESS = Mid(BUSINESS, 1, 32)
                liBusinessGID = GetBusinessID(BUSINESS)
                BUSINESS = FormatTextInput(BUSINESS)

                OFFICE_MAIL_ID = loMigrationData.Item("OFFICE_MAIL_ID").ToString.Trim
                OFFICE_MAIL_ID = FormatTextInput(OFFICE_MAIL_ID)

                PERSONAL_EMAIL_ID = loMigrationData.Item("PERSONAL_EMAIL_ID").ToString.Trim
                PERSONAL_EMAIL_ID = FormatTextInput(PERSONAL_EMAIL_ID)

                CONTACT_NO = loMigrationData.Item("CONTACT_NO").ToString.Trim
                MOBILE_NO = loMigrationData.Item("MOBILE_NO").ToString.Trim
                EMP_ACCOUNTNO = loMigrationData.Item("EMP_ACCOUNTNO").ToString.Trim

                EMP_BANKNAME = loMigrationData.Item("EMP_BANKNAME").ToString.Trim
                liBankGID = GetBankGID(EMP_BANKNAME)
                EMP_BANKNAME = Mid(EMP_BANKNAME, 1, 8)

                EMP_IFS_NO = loMigrationData.Item("EMP_IFS_NO").ToString.Trim
                BRANCH_CODE = loMigrationData.Item("BRANCH_CODE").ToString.Trim

                BSCC_NAME = loMigrationData.Item("BSCC_NAME").ToString.Trim
                liFCCCGID = GetBSCC(BSCC_NAME)
                BSCC_NAME = Mid(BSCC_NAME, 1, 16)
                BSCC_NAME = FormatTextInput(BSCC_NAME)

                EMP_REPORTINGTO = loMigrationData.Item("EMP_REPORTINGTO").ToString.Trim
                DATE_OF_RESIGNATION = loMigrationData.Item("DATE_OF_RESIGNATION")
                EMPLOYEE_STATUS = loMigrationData.Item("EMPLOYEE_STATUS").ToString.Trim

                PRODUCT = loMigrationData.Item("PRODUCT").ToString.Trim
                liProductGID = GetProductGID(PRODUCT)
                PRODUCT = Mid(PRODUCT, 1, 16)
                PRODUCT = FormatTextInput(PRODUCT)

                OU_NAME = loMigrationData.Item("OU_NAME").ToString.Trim
                liOUGID = GetOUGID(OU_NAME)
                OU_NAME = FormatTextInput(OU_NAME)

                BRANCH_CODE = GetBranchID(Mid(BRANCH_CODE, 1, 8))

                If Not IsExistsAtIEM("EMPLOYEE", EMP_STAFFID) Then
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_mst_temployee(employee_code, employee_name, employee_dob, employee_gender, employee_addr1, "
                    lsSQL &= " employee_addr2, employee_addr3, employee_addr4, employee_city_name, employee_city_gid, "
                    lsSQL &= " employee_pincode, employee_doj, employee_grade_code, employee_grade_gid, employee_hris_designation, "
                    lsSQL &= " employee_iem_designation, employee_iem_designation_gid,employee_dept_name, employee_dept_gid, "
                    lsSQL &= " employee_unit_name, employee_unit_gid, employee_office_email, employee_personal_email,"
                    lsSQL &= " employee_contact_no, employee_mobile_no, employee_era_acc_no, employee_era_bank_code, employee_era_bank_gid,"
                    lsSQL &= " employee_era_ifsc_code, employee_fccc_code, employee_fccc_gid,employee_supervisor, "
                    lsSQL &= " employee_dor, employee_product_code, employee_product_gid, employee_ou_code, employee_ou_gid, "
                    lsSQL &= " employee_branch_code, employee_branch_gid, employee_status, employee_insert_by, employee_insert_date,employee_photo_flag,employee_photo_filename,employee_update_by, employee_update_date, HRIS_LASTMODIFIEDON) "
                    lsSQL &= " VALUES("

                    lsSQL &= "'" & EMP_STAFFID & "',"
                    lsSQL &= "'" & EMP_COMPANYNAME & "',"

                    If IsDate(DATE_OF_BIRTH) Then
                        lsSQL &= "'" & Format(CDate(DATE_OF_BIRTH), "yyyy-MM-dd") & "',"
                    Else
                        lsSQL &= "'1900-01-01',"
                    End If
                    lsSQL &= "'" & PER_GENDER & "',"
                    lsSQL &= "'" & CON_HOME_ADDRESS1 & "',"
                    lsSQL &= "'" & CON_HOME_ADDRESS2 & "',"
                    lsSQL &= "'" & CON_HOME_ADDRESS3 & "',"
                    lsSQL &= "'" & CON_HOME_ADDRESS4 & "',"

                    lsSQL &= "'" & CITY_NAME & "',"
                    lsSQL &= liCityID & ","
                    lsSQL &= "'" & PIN_CODE & "',"

                    lsSQL &= "'" & Format(CDate(DATE_OF_JOINING), "yyyy-MM-dd") & "',"


                    lsSQL &= "'" & GRADE_NAME & "',"
                    lsSQL &= liGradeGID & ","

                    lsSQL &= "'" & HRIS_DESIGNATION & "',"
                    lsSQL &= "'" & IEM_DESIGNATION & "',"
                    lsSQL &= liDesignationGID & ","

                    lsSQL &= "'" & DEPARTMENT & "',"
                    lsSQL &= liDepartmentGID & ","

                    lsSQL &= "'" & BUSINESS & "',"
                    lsSQL &= liBusinessGID & ","

                    lsSQL &= "'" & OFFICE_MAIL_ID & "',"
                    lsSQL &= "'" & PERSONAL_EMAIL_ID & "',"
                    lsSQL &= "'" & CONTACT_NO & "',"
                    lsSQL &= "'" & MOBILE_NO & "',"
                    lsSQL &= "'" & EMP_ACCOUNTNO & "',"

                    lsSQL &= "'" & EMP_BANKNAME & "',"
                    lsSQL &= liBankGID & ","

                    lsSQL &= "'" & EMP_IFS_NO & "',"

                    lsSQL &= "'" & BSCC_NAME & "',"
                    lsSQL &= liFCCCGID & ","


                    lsSQL &= "" & IIf(EMP_REPORTINGTO = "NULL", 0, Val(EMP_REPORTINGTO)) & ","

                    If IsDate(DATE_OF_RESIGNATION) Then
                        lsSQL &= "'" & Format(CDate(DATE_OF_RESIGNATION), "yyyy-MM-dd") & "',"
                    Else
                        lsSQL &= "Null,"
                    End If

                    lsSQL &= "'" & PRODUCT & "',"
                    lsSQL &= liProductGID & ","

                    lsSQL &= "'" & OU_NAME & "',"
                    lsSQL &= liOUGID & ","

                    lsSQL &= "'" & BRANCH_CODE & "',"
                    lsSQL &= liBranchGID & ","

                    lsSQL &= "'" & IIf(EMPLOYEE_STATUS = "Active", "Y", "N") & "',"

                    lsSQL &= "0,SYSDATETIME(),'N','',0,SYSDATETIME(),SYSDATETIME()) "

                    'employee_photo_flag, employee_photo_filename, employee_update_by, employee_update_date, HRIS_LASTMODIFIEDON

                    lbIsModified = False

                Else

                    lbIsModified = True

                End If

                If lsSQL <> "" Then _
                    lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                If lsSQL = "" Then
                    If lbIsModified Then
                        liModifiedCount += 1
                    Else
                        liNewCount += 1
                    End If
                Else
                    liErrored += 1
                End If

                Me.Text = "EMPLOYEE : NEW - " & liNewCount & ", MODIFIED - " & liModifiedCount & ", Errored - " & liErrored
                Application.DoEvents()

            End While

            SummaryInsert("EMPLOYEE", liNewCount, liModifiedCount, liErrored)

        End If
    End Sub

    Private Function GetBranchID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT branch_gid FROM iem_mst_tbranch "
        lsSQl &= " WHERE LOWER(branch_code)='" & psState.ToString.ToLower & "' "

        GetBranchID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        If GetBranchID = 0 Then
            lsSQl = ""
            lsSQl &= " INSERT INTO iem_mst_tbranch(branch_code,branch_name, branch_addr1, branch_addr2, "
            lsSQl &= " branch_addr3, branch_addr4, branch_city_name, branch_location_name,branch_flag, branch_branchtype_flag, "
            lsSQl &= " branch_branchtype_name, branch_incharge, branch_start_date, branch_insert_by, branch_insert_date, branch_active, branch_city_gid, branch_location_gid,branch_branchtype_gid) "
            lsSQl &= "  VALUES('" & psState & "','','','','','','','','','','',0,'1900-01-01',0,SYSDATETIME(),'Y',1,1,1) "

            lsSQl = loDBConnection.ExecuteNonQuerySQL(lsSQl)

            lsSQl = ""
            lsSQl &= " SELECT branch_gid FROM iem_mst_tbranch "
            lsSQl &= " WHERE LOWER(branch_code)='" & psState.ToString.ToLower & "' "

            GetBranchID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)
        End If

    End Function

    Private Function GetBankGID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT bank_gid FROM iem_mst_tbank "
        lsSQl &= " WHERE LOWER(bank_name)='" & psState.ToString.ToLower & "' "

        GetBankGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        If GetBankGID = 0 Then
            Dim lsBankCode As String
            Dim lsBankName As String

            lsBankCode = GetCode("iem_mst_tbank", "bank_name", psState, 8, "bank_code")
            lsBankName = psState

            lsSQl = ""
            lsSQl &= " INSERT INTO iem_mst_tbank(bank_code, bank_name, bank_insert_by, bank_insert_date, "
            lsSQl &= " HRIS_LASTMODIFIEDON ) VALUES('" & lsBankCode & "',"
            lsSQl &= "'" & lsBankName & "',0,SYSDATETIME(),SYSDATETIME()) "

            lsSQl = ""
            lsSQl &= " SELECT bank_gid FROM iem_mst_tbank "
            lsSQl &= " WHERE LOWER(bank_name)='" & psState.ToString.ToLower & "' "

            GetBankGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)
        End If
    End Function

    Private Function GetBSCC(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String

        lsSQl = ""
        lsSQl &= " SELECT fccc_gid FROM iem_mst_tfccc "
        lsSQl &= " WHERE LOWER(fccc_name)='" & psState.ToString.ToLower & "' "

        GetBSCC = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        If GetBSCC = 0 Then

            Dim lsFCCCode = GetCode("iem_mst_tfccc", "fccc_name", psState, 16, "fccc_code")

            lsSQl = ""
            lsSQl &= " INSERT INTO iem_mst_tfccc(fccc_code,fccc_name, fccc_insert_by, fccc_insert_date, "
            lsSQl &= " HRIS_LASTMODIFIEDON, fccc_cc_gid,fccc_fc_gid,fccc_cc_code,fccc_fc_code) VALUES('" & lsFCCCode & "','" & psState & "',0,SYSDATETIME(),SYSDATETIME(),0,0,'','') "

            lsSQl = loDBConnection.ExecuteNonQuerySQL(lsSQl)

            lsSQl = ""
            lsSQl &= " SELECT fccc_gid FROM iem_mst_tfccc "
            lsSQl &= " WHERE LOWER(fccc_name)='" & psState.ToString.ToLower & "' "

            GetBSCC = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)
        End If

    End Function

    Private Function GetProductGID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)


        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT product_gid FROM iem_mst_tproduct "
        lsSQl &= " WHERE LOWER(product_name)='" & psState.ToString.ToLower & "' "

        GetProductGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        If GetProductGID = 0 Then

            Dim lsProductCode As String

            lsProductCode = GetCode("iem_mst_tproduct", "product_name", psState, 16, "product_code")

            lsSQl = ""
            lsSQl &= " INSERT INTO iem_mst_tproduct(product_code,product_name, product_insert_by, product_insert_date, "
            lsSQl &= " HRIS_LASTMODIFIEDON ) VALUES('" & lsProductCode & "','" & psState & "',0,SYSDATETIME(),SYSDATETIME()) "

            lsSQl = loDBConnection.ExecuteNonQuerySQL(lsSQl)

            lsSQl = ""
            lsSQl &= " SELECT product_gid FROM iem_mst_tproduct "
            lsSQl &= " WHERE LOWER(product_name)='" & psState.ToString.ToLower & "' "

            GetProductGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)
        End If
    End Function

    Private Function GetOUGID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT ou_gid FROM iem_mst_tou "
        lsSQl &= " WHERE LOWER(ou_name)='" & psState.ToString.ToLower & "' "

        GetOUGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        If GetOUGID = 0 Then
            lsSQl = ""
            lsSQl &= " INSERT INTO iem_mst_tou(ou_name, ou_insert_by, ou_insert_date, "
            lsSQl &= " HRIS_LASTMODIFIEDON ) VALUES('" & psState & "',0,SYSDATETIME(),SYSDATETIME()) "

            lsSQl = loDBConnection.ExecuteNonQuerySQL(lsSQl)

            lsSQl = ""
            lsSQl &= " SELECT ou_gid FROM iem_mst_tou "
            lsSQl &= " WHERE LOWER(ou_name)='" & psState.ToString.ToLower & "' "

            GetOUGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)
        End If
    End Function

    Private Function GetBusinessID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT unit_gid FROM iem_mst_tunit "
        lsSQl &= " WHERE LOWER(unit_name)='" & psState.ToString.ToLower & "' "

        GetBusinessID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        If GetBusinessID = 0 Then

            Dim lsUnitCode As String
            Dim lsUnitName As String

            lsUnitName = psState
            lsUnitCode = GetCode("iem_mst_tunit", "unit_name", lsUnitName, 8, "unit_code")

            lsSQl = ""
            lsSQl &= " INSERT INTO iem_mst_tunit(unit_code,unit_name, unit_insert_by, unit_insert_date) "
            lsSQl &= " VALUES('" & lsUnitCode & "','" & lsUnitName & "',0,SYSDATETIME()) "

            lsSQl = loDBConnection.ExecuteNonQuerySQL(lsSQl)

            lsSQl = ""
            lsSQl &= " SELECT unit_gid FROM iem_mst_tunit "
            lsSQl &= " WHERE LOWER(unit_name)='" & psState.ToString.ToLower & "' "

            GetBusinessID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)
        End If
    End Function

    Private Function GetDesignationGID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT designation_gid FROM iem_mst_tdesignation "
        lsSQl &= " WHERE LOWER(designation_name)='" & psState.ToString.ToLower & "' "

        GetDesignationGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        If GetDesignationGID = 0 Then
            Dim lsDesignationName As String
            Dim lsDesignationCode As String

            lsDesignationName = psState
            lsDesignationCode = GetCode("iem_mst_tdesignation", "designation_name", lsDesignationName, 8, "designation_code")

            lsSQl = ""
            lsSQl &= " INSERT INTO iem_mst_tdesignation(designation_code,designation_name, designation_level, designation_insert_by, designation_insert_date, "
            lsSQl &= " HRIS_LASTMODIFIEDON ) VALUES('" & lsDesignationCode & "','" & psState & "',1,0,SYSDATETIME(),SYSDATETIME()) "

            lsSQl = loDBConnection.ExecuteNonQuerySQL(lsSQl)


            lsSQl = ""
            lsSQl &= " SELECT designation_gid FROM iem_mst_tdesignation "
            lsSQl &= " WHERE LOWER(designation_name)='" & psState.ToString.ToLower & "' "

            ' lsSQl = loDBConnection.ExecuteNonQuerySQL(lsSQl)

            GetDesignationGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)
        End If
    End Function


    Private Function GetDepartmentGID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT dept_gid FROM iem_mst_tdept "
        lsSQl &= " WHERE LOWER(dept_name)='" & psState.ToString.ToLower & "' "

        GetDepartmentGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        If GetDepartmentGID = 0 Then

            Dim lsDeptCode As String
            lsDeptCode = GetCode("iem_mst_tdept", "dept_name", psState, 8, "dept_code")

            lsSQl = ""
            lsSQl &= " INSERT INTO iem_mst_tdept(dept_code,dept_name, dept_insert_by, dept_insert_date, "
            lsSQl &= " HRIS_LASTMODIFIEDON ) VALUES('" & lsDeptCode & "','" & psState & "',"
            lsSQl &= " 0,SYSDATETIME(),SYSDATETIME()) "

            lsSQl = loDBConnection.ExecuteNonQuerySQL(lsSQl)

            lsSQl = ""
            lsSQl &= " SELECT dept_gid FROM iem_mst_tdept "
            lsSQl &= " WHERE LOWER(dept_name)='" & psState.ToString.ToLower & "' "

            GetDepartmentGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)
        End If
    End Function

    Private Function FormatTextInput(ByVal psInput As String) As String
        psInput = psInput.Replace("'", "''")
        psInput = psInput.Replace("\", "\\")
        FormatTextInput = psInput
    End Function

    Private Function GetGradeGID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT grade_gid FROM iem_mst_tgrade "
        lsSQl &= " WHERE LOWER(grade_code)='" & psState.ToString.ToLower & "' "

        GetGradeGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        If GetGradeGID = 0 Then

            Dim lsGradeCode As String
            lsGradeCode = GetCode("iem_mst_tgrade", "grade_name", psState, 8, "grade_code")

            lsSQl = ""
            lsSQl &= " INSERT INTO iem_mst_tgrade(grade_code, grade_name, grade_level, grade_insert_by, grade_insert_date, "
            lsSQl &= " HRIS_LASTMODIFIEDON ) VALUES('" & lsGradeCode & "','" & psState & "',"
            lsSQl &= " '0',0,SYSDATETIME(),SYSDATETIME()) "

            lsSQl = loDBConnection.ExecuteNonQuerySQL(lsSQl)

            lsSQl = ""
            lsSQl &= " SELECT grade_gid FROM iem_mst_tgrade "
            lsSQl &= " WHERE LOWER(grade_code)='" & lsGradeCode & "' "

            GetGradeGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)
        End If
    End Function


    Private Function GetRegionGID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT region_gid FROM iem_mst_tregion "
        lsSQl &= " WHERE LOWER(region_name)='" & psState.ToString.ToLower & "' "

        GetRegionGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        If GetRegionGID = 0 Then
            lsSQl = ""
            lsSQl &= " INSERT INTO iem_mst_tregion(region_name, region_insert_by, region_insert_date, "
            lsSQl &= " HRIS_LASTMODIFIEDON ) VALUES('" & psState & "',0,SYSDATETIME(),SYSDATETIME()) "

            loDBConnection.ExecuteNonQuerySQL(lsSQl)

            lsSQl = ""
            lsSQl &= " SELECT region_gid FROM iem_mst_tregion "
            lsSQl &= " WHERE LOWER(region_name)='" & psState.ToString.ToLower & "' "

            GetRegionGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)
        End If
    End Function

    Private Function GetCountryGID(ByVal psState As String)

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT country_gid FROM iem_mst_tcountry "
        lsSQl &= " WHERE LOWER(country_name)='" & psState.ToString.ToLower & "' "

        GetCountryGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        If GetCountryGID = 0 Then

            Dim lsCountryCode As String
            lsCountryCode = GetCode("iem_mst_tcountry", "country_name", psState, 8, "country_code")

            lsSQl = ""
            lsSQl &= " INSERT INTO iem_mst_tcountry(country_code,country_name, country_currency_gid, country_currency_code, country_insert_by, country_insert_date, "
            lsSQl &= " HRIS_LASTMODIFIEDON ) VALUES('" & lsCountryCode & "','" & psState & "',1, 'IND',0,"
            lsSQl &= " SYSDATETIME(),SYSDATETIME()) "

            loDBConnection.ExecuteNonQuerySQL(lsSQl)

            lsSQl = ""
            lsSQl &= " SELECT country_gid FROM iem_mst_tcountry "
            lsSQl &= " WHERE LOWER(country_name)='" & psState.ToString.ToLower & "' "

            GetCountryGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        End If
    End Function

    Private Function GetStateGID(ByVal psState As String, ByVal piRegionID As Integer, ByVal piCountryID As Integer)

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT state_gid FROM iem_mst_tstate "
        lsSQl &= " WHERE LOWER(state_name)='" & psState.ToString.ToLower & "' "

        GetStateGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        If GetStateGID = 0 Then

            Dim lsStateCode As String

            lsStateCode = GetCode("iem_mst_tstate", "state_name", psState, 8, "state_code")

            lsSQl = ""
            lsSQl &= " INSERT INTO iem_mst_tstate(state_code,state_name, state_region_gid, state_country_gid,state_insert_by, state_insert_date,state_region_name,state_country_code)  "
            lsSQl &= " VALUES('" & lsStateCode & "','" & psState & "'," & piRegionID & "," & piCountryID & ","
            lsSQl &= " 0,SYSDATETIME(),'','') "

            lsSQl = loDBConnection.ExecuteNonQuerySQL(lsSQl)

            lsSQl = ""
            lsSQl &= " SELECT state_gid FROM iem_mst_tstate "
            lsSQl &= " WHERE LOWER(state_name)='" & psState.ToString.ToLower & "' "

            GetStateGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        End If
    End Function

    Private Function GetCityGID(ByVal psState As String, ByVal piRegionID As Integer, ByVal piCountryID As Integer, ByVal piStateID As Integer)

        psState = FormatTextInput(psState)

        Dim lsSQl As String

        lsSQl = ""
        lsSQl &= " SELECT city_gid FROM iem_mst_tcity "
        lsSQl &= " WHERE LOWER(city_name)='" & psState.ToString.ToLower & "' "

        GetCityGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        If GetCityGID = 0 Then

            Dim lsCItyCode As String
            Dim lsCityName As String

            lsCityName = psState
            lsCItyCode = GetCode("iem_mst_tcity", "city_name", lsCityName, 8, "city_code")

            lsSQl = ""
            lsSQl &= " INSERT INTO iem_mst_tcity(city_code,city_name, city_region_gid, city_country_gid, city_state_gid,city_insert_by, city_insert_date,city_pincode,city_state_code,city_state_name,city_region_name,city_country_code,city_country_name,city_cityclass_code,city_tier_code)  "
            lsSQl &= " VALUES('" & lsCItyCode & "','" & lsCityName & "'," & piRegionID & "," & piCountryID & "," & piStateID & ","
            lsSQl &= " 0,SYSDATETIME(),'','','','','','','','') "

            lsSQl = loDBConnection.ExecuteNonQuerySQL(lsSQl)

            lsSQl = ""
            lsSQl &= " SELECT city_gid FROM iem_mst_tcity "
            lsSQl &= " WHERE LOWER(city_name)='" & psState.ToString.ToLower & "' "

            GetCityGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        End If
    End Function

    Private Sub FLIDS_HOLIDAY()
        Dim loMigrationData As SqlDataReader

        'Dim lsLastModifiedOn As String
        Dim lsTableName As String
        Dim lsSQL As String
        Dim liNewCount, liModifiedCount As Integer
        Dim liHolidayGID As Integer = 0
        Dim liRegionGID As Integer = 0

        Dim lsHolidayDescription As String
        Dim lsHolidayState As String
        Dim liErrored As Integer
        Dim lbIsModified As Boolean
        Dim ldHolidayDate As Date

        'lsSQL = ""
        'lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
        'lsSQL &= " FROM iem_mst_tcountry "

        'lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL)

        'lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)

        lsTableName = " [econnect].[dbo].[IEM_HOLIDAY_MASTER_FIC] "
        liNewCount = 0
        liModifiedCount = 0
        liErrored = 0
        lsSQL = ""
        lsSQL &= " SELECT * FROM " & lsTableName
        'lsSQL &= " WHERE LAST_MODIFIED_ON IS NULL "
        'lsSQL &= " OR LAST_MODIFIED_ON >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "

        loMigrationData = loDBConnection.GetDataReader(lsSQL)

        If loMigrationData.HasRows Then

            While loMigrationData.Read

                ldHolidayDate = loMigrationData.Item("HOLIDAY_DATE")
                lsHolidayDescription = loMigrationData.Item("HOLIDAY_DESCRIPTION")
                lsHolidayState = loMigrationData.Item("CALENDAR_NAME")

                If Not IsExistsAtIEM("HOLIDAY", ldHolidayDate) Then
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_mst_tholiday(holiday_date, holiday_name, holiday_national_flag, holiday_state_flag, holiday_cutoff_flag,"
                    lsSQL &= " holiday_insert_by, holiday_insert_date, HRIS_LASTMODIFIEDON ) "
                    lsSQL &= " VALUES('" & Format(CDate(ldHolidayDate), "yyyy-MM-dd hh:mm:ss") & "',"
                    lsSQL &= "'" & lsHolidayDescription & "','Y','Y','Y',0,SYSDATETIME(),SYSDATETIME()) "

                    lbIsModified = False

                Else
                    lsSQL = ""
                    lsSQL &= " UPDATE iem_mst_tholiday "
                    lsSQL &= " SET holiday_name='" & loMigrationData.Item("HOLIDAY_DESCRIPTION") & "',"
                    lsSQL &= " holiday_update_by=0,"
                    lsSQL &= " holiday_update_date=SYSDATETIME() "
                    'lsSQL &= " HRIS_LASTMODIFIEDON='" & loMigrationData.Item("LAST_MODIFIED_ON") & "' "
                    lsSQL &= " WHERE holiday_date='" & Format(CDate(loMigrationData.Item("HOLIDAY_DATE")), "yyyy-MM-dd hh:mm:ss") & "' "


                    lbIsModified = True
                End If

                lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                HolidayStateInsert(Format(CDate(ldHolidayDate), "yyyy-MM-dd hh:mm:ss"), loMigrationData.Item("CALENDAR_NAME"))

                If lsSQL = "" Then
                    If lbIsModified Then
                        liModifiedCount += 1
                    Else
                        liNewCount += 1
                    End If
                Else
                    liErrored += 1
                End If
            End While

            SummaryInsert("HOLIDAY", liNewCount, liModifiedCount, liErrored)

        End If
    End Sub

    '' ''Private Sub FLIDS_CITY()
    '' ''    Dim loMigrationData As SqlDataReader

    '' ''    'Dim lsLastModifiedOn As String
    '' ''    Dim lsTableName As String
    '' ''    Dim lsSQL As String
    '' ''    Dim liNewCount, liModifiedCount As Integer
    '' ''    Dim liHolidayGID As Integer = 0
    '' ''    Dim liRegionGID As Integer = 0


    '' ''    Dim lsBankCode As String
    '' ''    Dim lsBankName As String
    '' ''    Dim lbIsModified As Boolean

    '' ''    Dim liErroed As Integer


    '' ''    'lsSQL = ""
    '' ''    'lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
    '' ''    'lsSQL &= " FROM iem_mst_tcountry "

    '' ''    'lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL)

    '' ''    'lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)

    '' ''    lsTableName = " [econnect].[dbo].[City] "
    '' ''    liNewCount = 0
    '' ''    liModifiedCount = 0
    '' ''    liErroed = 0

    '' ''    lsSQL = ""
    '' ''    lsSQL &= " SELECT * FROM " & lsTableName
    '' ''    'lsSQL &= " WHERE LAST_MODIFIED_ON IS NULL "
    '' ''    'lsSQL &= " OR LAST_MODIFIED_ON >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "

    '' ''    loMigrationData = loDBConnection.GetDataReader(lsSQL)

    '' ''    If loMigrationData.HasRows Then

    '' ''        While loMigrationData.Read

    '' ''            If Not IsExistsAtIEM("CITY", loMigrationData.Item("CITY")) Then
    '' ''                lsSQL = ""
    '' ''                lsSQL &= " INSERT INTO iem_mst_tcity(city_code, city_name, city_country_gid, city_insert_by, city_insert_date, "
    '' ''                lsSQL &= " HRIS_LASTMODIFIEDON ) VALUES('" & Mid(loMigrationData.Item("CODE"), 1, 8) & "',"
    '' ''                lsSQL &= "'" & loMigrationData.Item("CITY") & "'," & GetCountryGID(loMigrationData.Item("COUNTY")) & ",SYSDATETIME(),SYSDATETIME()) "

    '' ''                lbIsModified = False

    '' ''            Else
    '' ''                lsSQL = ""
    '' ''                lsSQL &= " UPDATE iem_mst_tcity "
    '' ''                lsSQL &= " SET city_code='" & loMigrationData.Item("CODE") & "',"
    '' ''                lsSQL &= " city_update_by=0,"
    '' ''                lsSQL &= " city_update_date=SYSDATETIME(),"
    '' ''                lsSQL &= " city_country_gid = " & GetCountryGID(loMigrationData.Item("COUNTRY"))
    '' ''                'lsSQL &= " HRIS_LASTMODIFIEDON='" & loMigrationData.Item("LAST_MODIFIED_ON") & "' "
    '' ''                lsSQL &= " WHERE LOWER(city_name)='" & loMigrationData.Item("CITY").ToString.ToLower & "' "


    '' ''                lbIsModified = True
    '' ''            End If

    '' ''            lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

    '' ''        End While

    '' ''        lsSQL = ""
    '' ''        lsSQL &= " INSERT INTO iem_mig_tflids(FLIDS_DATE, FLIDS_UPDATEAT, FLIDS_NEWINSERT, FLIDS_MODIFIED) "
    '' ''        lsSQL &= " VALUES(SYSDATETIME(),'STATE'," & liNewCount & "," & liModifiedCount & ")"

    '' ''        loDBConnection.ExecuteNonQuerySQL(lsSQL)

    '' ''    End If
    '' ''End Sub

    '' ''Private Sub FLIDS_STATE()
    '' ''    Dim loMigrationData As SqlDataReader

    '' ''    'Dim lsLastModifiedOn As String
    '' ''    Dim lsTableName As String
    '' ''    Dim lsSQL As String
    '' ''    Dim liNewCount, liModifiedCount As Integer
    '' ''    Dim liHolidayGID As Integer = 0
    '' ''    Dim liRegionGID As Integer = 0


    '' ''    'lsSQL = ""
    '' ''    'lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
    '' ''    'lsSQL &= " FROM iem_mst_tcountry "

    '' ''    'lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL)

    '' ''    'lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)

    '' ''    lsTableName = " [econnect].[dbo].[State] "
    '' ''    liNewCount = 0
    '' ''    liModifiedCount = 0

    '' ''    lsSQL = ""
    '' ''    lsSQL &= " SELECT * FROM " & lsTableName
    '' ''    'lsSQL &= " WHERE LAST_MODIFIED_ON IS NULL "
    '' ''    'lsSQL &= " OR LAST_MODIFIED_ON >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "

    '' ''    loMigrationData = loDBConnection.GetDataReader(lsSQL)

    '' ''    If loMigrationData.HasRows Then

    '' ''        While loMigrationData.Read

    '' ''            If Not IsExistsAtIEM("STATE", loMigrationData.Item("STATE")) Then
    '' ''                lsSQL = ""
    '' ''                lsSQL &= " INSERT INTO iem_mst_tstate(state_code, state_name, state_country_gid, state_insert_by, state_insert_date, "
    '' ''                lsSQL &= " HRIS_LASTMODIFIEDON ) VALUES('" & Mid(loMigrationData.Item("CODE"), 1, 8) & "',"
    '' ''                lsSQL &= "'" & loMigrationData.Item("STATE") & "'," & GetCountryGID(loMigrationData.Item("COUNTY")) & ",SYSDATETIME(),SYSDATETIME()) "

    '' ''                liNewCount += 1

    '' ''            Else
    '' ''                lsSQL = ""
    '' ''                lsSQL &= " UPDATE iem_mst_tstate "
    '' ''                lsSQL &= " SET state_code='" & loMigrationData.Item("CODE") & "',"
    '' ''                lsSQL &= " state_update_by=0,"
    '' ''                lsSQL &= " state_update_date=SYSDATETIME(),"
    '' ''                lsSQL &= " state_country_gid = " & GetCountryGID(loMigrationData.Item("COUNTRY"))
    '' ''                'lsSQL &= " HRIS_LASTMODIFIEDON='" & loMigrationData.Item("LAST_MODIFIED_ON") & "' "
    '' ''                lsSQL &= " WHERE LOWER(state_name)='" & loMigrationData.Item("STATE").ToString.ToLower & "' "


    '' ''                liModifiedCount += 1
    '' ''            End If

    '' ''            lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

    '' ''        End While

    '' ''        lsSQL = ""
    '' ''        lsSQL &= " INSERT INTO iem_mig_tflids(FLIDS_DATE, FLIDS_UPDATEAT, FLIDS_NEWINSERT, FLIDS_MODIFIED) "
    '' ''        lsSQL &= " VALUES(SYSDATETIME(),'STATE'," & liNewCount & "," & liModifiedCount & ")"

    '' ''        loDBConnection.ExecuteNonQuerySQL(lsSQL)

    '' ''    End If
    '' ''End Sub



    Private Function HolidayStateInsert(ByVal pdHolidayDate As Date, ByVal psState As String) As Boolean

        psState = FormatTextInput(psState)

        Dim lsSQL As String
        Dim liHolidayGID As Integer
        Dim liStateGID As Integer
        Dim liHolidayStateGID As Integer

        HolidayStateInsert = True

        lsSQL = ""
        lsSQL &= " SELECT MAX(holiday_gid) FROM iem_mst_tholiday "
        lsSQL &= " WHERE holiday_date='" & Format(CDate(pdHolidayDate), "yyyy-MM-dd hh:mm:ss") & "' "

        liHolidayGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

        lsSQL = ""
        lsSQL &= " SELECT state_gid FROM iem_mst_tstate "
        lsSQL &= " WHERE state_name='" & psState & "' "

        liStateGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

        lsSQL = ""
        lsSQL &= " SELECT holidaystate_gid "
        lsSQL &= " FROM iem_mst_tholidaystate "
        lsSQL &= " WHERE holidaystate_holiday_gid = " & liHolidayGID & "  "
        lsSQL &= " AND holidaystate_state_gid = " & liStateGID & " "

        liHolidayStateGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

        If liHolidayGID > 0 And liStateGID > 0 And liHolidayStateGID = 0 Then
            lsSQL = ""
            lsSQL &= " INSERT INTO iem_mst_tholidaystate(holidaystate_holiday_gid, holidaystate_state_gid) "
            lsSQL &= " VALUES(" & liHolidayGID & "," & liStateGID & ") "

            lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

        Else

            HolidayStateInsert = False
        End If


    End Function

    Private Sub FLIDS_DESIGNATION()
        Dim loMigrationData As SqlDataReader

        Dim lsLastModifiedOn As String
        Dim lsTableName As String
        Dim lsSQL As String
        Dim liNewCount, liModifiedCount, liErrored As Integer
        Dim lbIsModified As Boolean

        Dim lsDesignationCode As String
        Dim lsDesignationName As String
        Dim lsDesignationLevel As String


        lsSQL = ""
        lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
        lsSQL &= " FROM iem_mst_tdesignation "

        lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL)

        lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)

        lsTableName = " [econnect].[dbo].[IEM_DESIGNATION_FIELDS_FIC] "
        liNewCount = 0
        liModifiedCount = 0
        liErrored = 0

        lsSQL = ""
        lsSQL &= " SELECT * FROM " & lsTableName
        lsSQL &= " WHERE LAST_MODIFIED_ON IS NULL "
        lsSQL &= " OR LAST_MODIFIED_ON >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "

        loMigrationData = loDBConnection.GetDataReader(lsSQL)

        If loMigrationData.HasRows Then

            While loMigrationData.Read

                lsDesignationCode = loMigrationData.Item("DESIGNATION_NAME").ToString.Trim
                lsDesignationName = loMigrationData.Item("DESIGNATION_DESCRIPTION").ToString.Trim

                lsDesignationCode = GetSuitableDesignationCode(lsDesignationCode, lsDesignationName)
                lsDesignationLevel = loMigrationData.Item("DESIGNATION_LEVEL").ToString.Trim

                If Not IsExistsAtIEM("DESIGNATION", lsDesignationName) Then
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_mst_tdesignation(designation_code, designation_name, designation_level, designation_insert_by, designation_insert_date, "
                    lsSQL &= " HRIS_LASTMODIFIEDON ) VALUES('" & lsDesignationCode & "','" & lsDesignationName & "',"
                    lsSQL &= "'" & lsDesignationLevel & "',0,SYSDATETIME(),SYSDATETIME()) "

                    lbIsModified = False

                Else
                    lsSQL = ""
                    lsSQL &= " UPDATE iem_mst_tdesignation "
                    lsSQL &= " SET designation_name='" & Mid(lsDesignationName, 1, 32) & "',"
                    lsSQL &= " designation_level=" & lsDesignationLevel & ","
                    lsSQL &= " designation_update_by=0,"
                    lsSQL &= " designation_update_date=SYSDATETIME(),"
                    lsSQL &= " HRIS_LASTMODIFIEDON='" & loMigrationData.Item("LAST_MODIFIED_ON") & "' "
                    lsSQL &= " WHERE LOWER(designation_code)='" & lsDesignationCode.ToLower & "'"

                    lbIsModified = True
                End If

                lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                If lsSQL = "" Then
                    If lbIsModified Then
                        liModifiedCount += 1
                    Else
                        liNewCount += 1
                    End If
                Else
                    liErrored += 1
                End If
            End While

            SummaryInsert("DESIGNATION", liNewCount, liModifiedCount, liErrored)

        End If
    End Sub

    Private Function GetSuitableDesignationCode(ByVal psCode As String, ByVal psName As String)
        Dim lsSQL As String

        lsSQL = ""
        lsSQL &= " SELECT designation_code FROM iem_mst_tdesignation "
        lsSQL &= " WHERE LOWER(designation_name)='" & psName.ToString.ToLower & "' "

        GetSuitableDesignationCode = loDBConnection.GetExecuteScalar(lsSQL).ToString

        If GetSuitableDesignationCode.ToString.Trim = "" Then
            psCode = Mid(psCode, 1, InStr(psCode, " "))
            psCode = Mid(psCode, 1, 6).Trim

            lsSQL = ""
            lsSQL &= " SELECT COUNT(*) FROM iem_mst_tdesignation "
            lsSQL &= " WHERE designation_code LIKE '" & psCode & "%' "

            Dim liCount As Integer
            liCount = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

            psCode = psCode & Format(liCount + 1, "00")

            GetSuitableDesignationCode = psCode
        End If

    End Function

    Private Function GetCode(ByVal psTableName As String, ByVal psFieldName As String, ByVal psValueToCheck As String, ByVal piFieldLength As Integer, ByVal psCodeFieldName As String)

        Dim lsSQL As String

        If InStr(psValueToCheck, " ") > 0 Then _
            psValueToCheck = Mid(psValueToCheck, 1, InStr(psValueToCheck, " "))

        psValueToCheck = Mid(psValueToCheck, 1, piFieldLength - 2).Trim

        lsSQL = ""
        lsSQL &= " SELECT COUNT(*) FROM " & psTableName
        lsSQL &= " WHERE " & psCodeFieldName & " LIKE '" & psValueToCheck & "%' "

        Dim liCount As Integer
        liCount = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)

        GetCode = psValueToCheck & Format(liCount + 1, "00")
    End Function

    Private Sub FLIDS_DEPARTMENT()
        Dim loMigrationData As SqlDataReader

        Dim lsLastModifiedOn As String
        Dim lsTableName As String
        Dim lsSQL As String
        Dim liNewCount, liModifiedCount, liErrored As Integer

        Dim lsDeptCode As String
        Dim lsDeptName As String

        Dim lbIsModified As Boolean


        lsSQL = ""
        lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
        lsSQL &= " FROM iem_mst_tdept "

        lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL)

        lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)

        lsTableName = " [econnect].[dbo].[IEM_DEPARTMENT_FIELDS_FIC] "
        liNewCount = 0
        liModifiedCount = 0

        lsSQL = ""
        lsSQL &= " SELECT * FROM " & lsTableName
        lsSQL &= " WHERE LAST_MODIFIED_ON IS NULL "
        lsSQL &= " OR LAST_MODIFIED_ON >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "

        loMigrationData = loDBConnection.GetDataReader(lsSQL)

        If loMigrationData.HasRows Then

            While loMigrationData.Read

                lsDeptCode = loMigrationData.Item("DEPARTMENT_CODE")
                lsDeptCode = FormatTextInput(lsDeptCode)

                lsDeptName = loMigrationData.Item("DEPARTMENT_NAME")
                lsDeptName = FormatTextInput(lsDeptName)
                lsDeptName = Mid(lsDeptName, 1, 32)
                lsDeptCode = Mid(lsDeptCode, 1, 8)

                If Not IsExistsAtIEM("DEPARTMENT", lsDeptCode) Then
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_mst_tdept(dept_code, dept_name, dept_insert_by, dept_insert_date, "
                    lsSQL &= " HRIS_LASTMODIFIEDON ) VALUES('" & lsDeptCode & "','" & lsDeptName & "',"
                    lsSQL &= " 0,SYSDATETIME(),SYSDATETIME()) "

                    lbIsModified = False

                Else
                    lsSQL = ""
                    lsSQL &= " UPDATE iem_mst_tdept "
                    lsSQL &= " SET dept_name='" & lsDeptName & "',"
                    lsSQL &= " dept_update_by=0,"
                    lsSQL &= " dept_update_date=SYSDATETIME(),"
                    If IsDate(loMigrationData.Item("LAST_MODIFIED_ON")) Then
                        lsSQL &= " HRIS_LASTMODIFIEDON='" & loMigrationData.Item("LAST_MODIFIED_ON") & "' "
                    Else
                        lsSQL &= " HRIS_LASTMODIFIEDON=Null "
                    End If
                    lsSQL &= " WHERE LOWER(dept_code)='" & lsDeptCode.ToLower & "'"

                    lbIsModified = True
                End If

                lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                If lsSQL = "" Then
                    If lbIsModified Then
                        liModifiedCount += 1
                    Else
                        liNewCount += 1
                    End If
                Else
                    liErrored += 1
                End If

            End While

            SummaryInsert("DEPARTMENT", liNewCount, liModifiedCount, liErrored)

        End If
    End Sub


    Private Sub FLIDS_GRADE()
        Dim loMigrationData As SqlDataReader

        Dim lsLastModifiedOn As String
        Dim lsTableName As String
        Dim lsSQL As String
        Dim liNewCount, liModifiedCount, liErrored As Integer
        Dim lbIsModified As Boolean

        Dim lsGradeCode As String
        Dim lsGradename As String
        Dim lsGradeLevel As String

        lsSQL = ""
        lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
        lsSQL &= " FROM iem_mst_tgrade "

        lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL)

        lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)

        lsTableName = " [econnect].[dbo].[IEM_GRADE_FIELDS_FIC] "
        liNewCount = 0
        liModifiedCount = 0

        lsSQL = ""
        lsSQL &= " SELECT * FROM " & lsTableName
        lsSQL &= " WHERE LAST_MODIFIED_ON IS NULL "
        lsSQL &= " OR LAST_MODIFIED_ON >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "

        loMigrationData = loDBConnection.GetDataReader(lsSQL)

        If loMigrationData.HasRows Then

            While loMigrationData.Read

                lsGradeCode = loMigrationData.Item("GRADE_NAME").ToString.Trim
                lsGradeCode = FormatTextInput(lsGradeCode)


                lsGradename = loMigrationData.Item("GRADE_DESCRIPTION").ToString.Trim
                lsGradename = FormatTextInput(lsGradename)


                lsGradeCode = GetCode("iem_mst_tgrade", "grade_name", lsGradeCode, 8, "grade_code")


                lsGradeLevel = loMigrationData.Item("GRADE_HIERARCHY").ToString.Trim
                lsGradeLevel = FormatTextInput(lsGradeLevel)

                If Not IsExistsAtIEM("GRADE", lsGradeCode) Then
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_mst_tgrade(grade_code, grade_name, grade_level, grade_insert_by, grade_insert_date, "
                    lsSQL &= " HRIS_LASTMODIFIEDON ) VALUES('" & lsGradeCode & "','" & lsGradename & "',"
                    lsSQL &= "'" & lsGradeLevel & "', 0,SYSDATETIME(),SYSDATETIME()) "

                    lbIsModified = False

                Else
                    lsSQL = ""
                    lsSQL &= " UPDATE iem_mst_tgrade "
                    lsSQL &= " SET grade_name='" & lsGradename & "',"
                    lsSQL &= " grade_update_by=0,"
                    lsSQL &= " grade_update_date=SYSDATETIME(),"
                    lsSQL &= " HRIS_LASTMODIFIEDON='" & loMigrationData.Item("LAST_MODIFIED_ON") & "' "
                    lsSQL &= " WHERE LOWER(grade_code)='" & lsGradeCode.ToLower & "'"

                    lbIsModified = True
                End If

                lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)
                If lsSQL = "" Then
                    If lbIsModified Then
                        liModifiedCount += 1
                    Else
                        liNewCount += 1
                    End If
                Else
                    liErrored += 1
                End If
            End While

            SummaryInsert("GRADE", liNewCount, liModifiedCount, liErrored)

        End If
    End Sub


    Private Sub FLIDS_COUNTRY()
        Dim loMigrationData As SqlDataReader

        Dim lsLastModifiedOn As String
        Dim lsTableName As String
        Dim lsSQL As String
        Dim liNewCount, liModifiedCount, liErrored As Integer
        Dim lbIsModified As Boolean

        Dim lsCountryCode As String
        Dim lsCountryName As String

        Dim liCurrencyGID As Integer = 1
        Dim lsCurrencyCode As String = "IND"

        lsSQL = ""
        lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
        lsSQL &= " FROM iem_mst_tcountry "

        lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL)

        lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)

        lsTableName = " [econnect].[dbo].[IEM_COUNTRY_FIELDS_FIC] "
        liNewCount = 0
        liModifiedCount = 0
        liErrored = 0

        lsSQL = ""
        lsSQL &= " SELECT * FROM " & lsTableName
        lsSQL &= " WHERE LAST_MODIFIED_ON IS NULL "
        lsSQL &= " OR LAST_MODIFIED_ON >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "

        loMigrationData = loDBConnection.GetDataReader(lsSQL)

        If loMigrationData.HasRows Then

            While loMigrationData.Read

                lsCountryCode = loMigrationData.Item("COUNTRY_CODE").ToString.Trim
                lsCountryCode = FormatTextInput(lsCountryCode)

                lsCountryName = loMigrationData.Item("COUNTRY_NAME").ToString.Trim
                lsCountryName = FormatTextInput(lsCountryName)

                lsCountryCode = GetCode("iem_mst_tcountry", "country_name", lsCountryName, 8, "country_code")


                If Not IsExistsAtIEM("COUNTRY", lsCountryCode) Then
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_mst_tcountry(country_code, country_name, country_insert_by, country_insert_date, "
                    lsSQL &= " country_currency_gid, country_currency_code,HRIS_LASTMODIFIEDON ) VALUES('" & lsCountryCode & "','" & lsCountryName & "',"
                    lsSQL &= " 0,SYSDATETIME()," & liCurrencyGID & ",'" & lsCountryCode & "',SYSDATETIME()) "

                    lbIsModified = False
                Else
                    lsSQL = ""
                    lsSQL &= " UPDATE iem_mst_tcountry "
                    lsSQL &= " SET country_name='" & lsCountryName & "',"
                    lsSQL &= " country_update_by=0,"
                    lsSQL &= " country_update_date=SYSDATETIME(),"


                    If IsDate(loMigrationData.Item("LAST_MODIFIED_ON")) Then
                        lsSQL &= " HRIS_LASTMODIFIEDON='" & loMigrationData.Item("LAST_MODIFIED_ON") & "' "
                    Else
                        lsSQL &= " HRIS_LASTMODIFIEDON=Null "
                    End If

                    'lsSQL &= " HRIS_LASTMODIFIEDON='" & loMigrationData.Item("LAST_MODIFIED_ON") & "' "
                    lsSQL &= " WHERE LOWER(country_code)='" & lsCountryCode.ToLower & "'"

                    lbIsModified = True
                End If

                lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                If lsSQL = "" Then
                    If lbIsModified Then
                        liModifiedCount += 1
                    Else
                        liNewCount += 1
                    End If
                Else
                    liErrored += 1
                End If

            End While

            SummaryInsert("COUNTRY", liNewCount, liModifiedCount, liErrored)
        End If
    End Sub

    Private Sub FLIDS_REGION()
        Dim loMigrationData As SqlDataReader

        Dim lsLastModifiedOn As String
        Dim lsTableName As String
        Dim lsSQL As String
        Dim liNewCount, liModifiedCount, liErrored As Integer

        Dim lsRegion As String
        Dim lbIsModified As Boolean


        lsSQL = ""
        lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
        lsSQL &= " FROM iem_mst_tregion "

        lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL)

        lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)

        lsTableName = " [econnect].[dbo].[IEM_REGION_FIELDS_FIC] "
        liNewCount = 0
        liModifiedCount = 0

        lsSQL = ""
        lsSQL &= " SELECT * FROM " & lsTableName
        lsSQL &= " WHERE LAST_MODIFIED_ON IS NULL "
        lsSQL &= " OR LAST_MODIFIED_ON >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "

        loMigrationData = loDBConnection.GetDataReader(lsSQL)

        If loMigrationData.HasRows Then

            While loMigrationData.Read

                lsRegion = loMigrationData.Item("REGION_NAME").ToString.Trim
                lsRegion = FormatTextInput(lsRegion)

                If Not IsExistsAtIEM("REGION", lsRegion) Then
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_mst_tregion(region_name, region_insert_by, region_insert_date, "
                    lsSQL &= " HRIS_LASTMODIFIEDON ) VALUES('" & lsRegion & "',0,SYSDATETIME(),SYSDATETIME()) "

                    lbIsModified = False

                Else
                    lsSQL = ""
                    lsSQL &= " UPDATE iem_mst_tregion "
                    lsSQL &= " SET region_name='" & lsRegion & "',"
                    lsSQL &= " region_update_by=0,"
                    lsSQL &= " region_update_date=SYSDATETIME(),"

                    If IsDate(loMigrationData.Item("LAST_MODIFIED_ON")) Then
                        lsSQL &= " HRIS_LASTMODIFIEDON='" & loMigrationData.Item("LAST_MODIFIED_ON") & "' "
                    Else
                        lsSQL &= " HRIS_LASTMODIFIEDON=Null "
                    End If

                    lsSQL &= " WHERE LOWER(region_name)='" & lsRegion.ToLower & "'"

                    lbIsModified = True
                End If

                lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                If lsSQL = "" Then
                    If lbIsModified Then
                        liModifiedCount += 1
                    Else
                        liNewCount += 1
                    End If
                Else
                    liErrored += 1
                End If

            End While

            SummaryInsert("REGION", liNewCount, liModifiedCount, liErrored)

        End If
    End Sub


    Private Sub FLIDS_PRODUCT()
        Dim loMigrationData As SqlDataReader

        Dim lsLastModifiedOn As String
        Dim lsTableName As String
        Dim lsSQL As String
        Dim liNewCount, liModifiedCount, liErrored As Integer

        Dim lsProduct As String
        Dim lsName As String
        Dim lbIsModified As Boolean


        lsSQL = ""
        lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
        lsSQL &= " FROM iem_mst_tproduct "

        lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL)

        lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)

        lsTableName = " [econnect].[dbo].[IEM_PRODUCT_FIELDS_FIC] "
        liNewCount = 0
        liModifiedCount = 0

        lsSQL = ""
        lsSQL &= " SELECT * FROM " & lsTableName
        lsSQL &= " WHERE MODIFIED_ON IS NULL "
        lsSQL &= " OR MODIFIED_ON >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "

        loMigrationData = loDBConnection.GetDataReader(lsSQL)

        If loMigrationData.HasRows Then

            While loMigrationData.Read

                lsProduct = loMigrationData.Item("MASTER3_ID").ToString.Trim
                lsProduct = FormatTextInput(lsProduct)

                lsName = loMigrationData.Item("MASTER3_NAME")
                lsName = FormatTextInput(lsName)


                If Not IsExistsAtIEM("PRODUCT", lsProduct) Then
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_mst_tproduct(product_code, product_name, product_insert_by, product_insert_date, "
                    lsSQL &= " HRIS_LASTMODIFIEDON ) VALUES('" & lsProduct & "',"
                    lsSQL &= "'" & lsName & "',0,SYSDATETIME(),SYSDATETIME()) "

                    lbIsModified = False

                Else
                    lsSQL = ""
                    lsSQL &= " UPDATE iem_mst_tproduct "
                    lsSQL &= " SET product_name='" & lsName & "',"
                    lsSQL &= " product_update_by=0,"
                    lsSQL &= " product_update_date=SYSDATETIME(),"
                    lsSQL &= " HRIS_LASTMODIFIEDON='" & loMigrationData.Item("MODIFIED_ON") & "' "
                    lsSQL &= " WHERE product_code='" & lsProduct & "'"

                    lbIsModified = True
                End If

                lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                If lsSQL = "" Then
                    If lbIsModified Then
                        liModifiedCount += 1
                    Else
                        liNewCount += 1
                    End If
                Else
                    liErrored += 1
                End If
            End While

            SummaryInsert("PRODUCT", liNewCount, liModifiedCount, liErrored)

        End If
    End Sub



    Private Sub FLIDS_OU()
        Dim loMigrationData As SqlDataReader

        Dim lsLastModifiedOn As String
        Dim lsTableName As String
        Dim lsSQL As String
        Dim liNewCount, liModifiedCount, liErrored As Integer

        Dim lsOUCode As String
        Dim lsOUName As String

        Dim lbIsModified As Boolean

        lsSQL = ""
        lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
        lsSQL &= " FROM iem_mst_tou "

        lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL)

        lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)

        lsTableName = " [econnect].[dbo].[IEM_OU_FIELDS_FIC] "
        liNewCount = 0
        liModifiedCount = 0

        lsSQL = ""
        lsSQL &= " SELECT * FROM " & lsTableName
        lsSQL &= " WHERE OU_UPDATED_ON IS NULL "
        lsSQL &= " OR OU_UPDATED_ON >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "

        loMigrationData = loDBConnection.GetDataReader(lsSQL)

        If loMigrationData.HasRows Then

            While loMigrationData.Read

                lsOUCode = loMigrationData.Item("OU_CODE")
                lsOUCode = FormatTextInput(lsOUCode)

                lsOUName = loMigrationData.Item("OU_NAME")
                lsOUName = FormatTextInput(lsOUName)

                If Not IsExistsAtIEM("OU", loMigrationData.Item("OU_CODE")) Then
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_mst_tou(ou_code, ou_name, ou_insert_by, ou_insert_date, "
                    lsSQL &= " HRIS_LASTMODIFIEDON ) VALUES('" & lsOUCode & "',"
                    lsSQL &= "'" & lsOUName & "',0,SYSDATETIME(),SYSDATETIME()) "

                    lbIsModified = False

                Else
                    lsSQL = ""
                    lsSQL &= " UPDATE iem_mst_tou "
                    lsSQL &= " SET ou_name='" & lsOUName & "',"
                    lsSQL &= " ou_update_by=0,"
                    lsSQL &= " ou_update_date=SYSDATETIME(),"
                    If IsDate(loMigrationData.Item("OU_UPDATED_ON")) Then
                        lsSQL &= " HRIS_LASTMODIFIEDON='" & loMigrationData.Item("OU_UPDATED_ON") & "' "
                    Else
                        lsSQL &= " HRIS_LASTMODIFIEDON=Null "
                    End If

                    lsSQL &= " WHERE ou_code='" & lsOUName & "'"

                    lbIsModified = True
                End If

                lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                If lsSQL = "" Then
                    If lbIsModified Then
                        liModifiedCount += 1
                    Else
                        liNewCount += 1
                    End If
                Else
                    liErrored += 1
                End If
            End While

            SummaryInsert("OU", liNewCount, liModifiedCount, liErrored)

        End If
    End Sub


    Private Sub FLIDS_FCCC()
        Dim loMigrationData As SqlDataReader

        Dim lsLastModifiedOn As String
        Dim lsTableName As String
        Dim lsSQL As String
        Dim liNewCount, liModifiedCount, liErrored As Integer

        Dim lbIsModified As Boolean

        Dim lsFCCC As String
        Dim lsFCCCName As String



        lsSQL = ""
        lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
        lsSQL &= " FROM iem_mst_tfccc "

        lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL)

        lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)

        lsTableName = " [econnect].[dbo].[IEM_BSCC_FIELDS_FIC] "
        liNewCount = 0
        liModifiedCount = 0

        lsSQL = ""
        lsSQL &= " SELECT * FROM " & lsTableName
        'lsSQL &= " WHERE LAST_MODIFIED_ON IS NULL "
        'lsSQL &= " OR LAST_MODIFIED_ON >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "

        loMigrationData = loDBConnection.GetDataReader(lsSQL)

        If loMigrationData.HasRows Then

            While loMigrationData.Read

                lsFCCC = loMigrationData.Item("BSCC_CODE").ToString.Trim
                lsFCCC = FormatTextInput(lsFCCC)

                lsFCCCName = loMigrationData.Item("BSCC_NAME").ToString.Trim
                lsFCCCName = FormatTextInput(lsFCCCName)

                If Not IsExistsAtIEM("FCCC", lsFCCC) Then
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_mst_tfccc(fccc_code, fccc_name, fccc_insert_by, fccc_insert_date, fccc_cc_gid,fccc_fc_gid,fccc_cc_code,fccc_fc_code, "
                    lsSQL &= " HRIS_LASTMODIFIEDON ) VALUES('" & lsFCCC & "',"
                    lsSQL &= "'" & lsFCCCName & "',0,SYSDATETIME(),0,0,'','', SYSDATETIME()) "

                    lbIsModified = False
                Else
                    lsSQL = ""
                    lsSQL &= " UPDATE iem_mst_tfccc "
                    lsSQL &= " SET fccc_name='" & loMigrationData.Item("BSCC_NAME") & "',"
                    lsSQL &= " fccc_update_by=0,"
                    lsSQL &= " fccc_update_date=SYSDATETIME(),"
                    'lsSQL &= " HRIS_LASTMODIFIEDON='" & loMigrationData.Item("LAST_MODIFIED_ON") & "' "
                    lsSQL &= " HRIS_LASTMODIFIEDON=SYSDATETIME() "
                    lsSQL &= " WHERE fccc_code='" & loMigrationData.Item("BSCC_CODE") & "'"

                    lbIsModified = True
                End If

                lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                If lsSQL = "" Then
                    If lbIsModified Then
                        liModifiedCount += 1
                    Else
                        liNewCount += 1
                    End If
                Else
                    liErrored += 1
                End If
            End While

            SummaryInsert("FCCC", liNewCount, liModifiedCount, liErrored)
        End If

    End Sub


    Private Sub FLIDS_BANK()
        Dim loMigrationData As SqlDataReader

        Dim lsLastModifiedOn As String
        Dim lsTableName As String
        Dim lsSQL As String
        Dim liNewCount, liModifiedCount, liErrored As Integer

        Dim lsBankCode As String
        Dim lsBankName As String
        Dim lbIsModified As Boolean


        lsSQL = ""
        lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
        lsSQL &= " FROM iem_mst_tbank "

        lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL)

        lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)

        lsTableName = " [econnect].[dbo].[IEM_BANK_FIELDS_FIC] "

        liNewCount = 0
        liModifiedCount = 0
        liErrored = 0

        lsSQL = ""
        lsSQL &= " SELECT * FROM " & lsTableName
        lsSQL &= " WHERE LAST_MODIFIED_ON IS NULL "
        lsSQL &= " OR LAST_MODIFIED_ON >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "

        loMigrationData = loDBConnection.GetDataReader(lsSQL)

        If loMigrationData.HasRows Then

            While loMigrationData.Read

                lsBankCode = loMigrationData.Item("BANK_CODE").ToString.Trim
                lsBankCode = FormatTextInput(lsBankCode)

                lsBankName = loMigrationData.Item("BANK_NAME").ToString.Trim
                lsBankName = FormatTextInput(lsBankCode)


                If Not IsExistsAtIEM("BANK", loMigrationData.Item("BANK_CODE")) Then
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_mst_tbank(bank_code, bank_name, bank_insert_by, bank_insert_date, "
                    lsSQL &= " HRIS_LASTMODIFIEDON ) VALUES('" & lsBankCode & "',"
                    lsSQL &= "'" & lsBankName & "',0,SYSDATETIME(),SYSDATETIME()) "

                    lbIsModified = False
                Else
                    lsSQL = ""
                    lsSQL &= " UPDATE iem_mst_tbank "
                    lsSQL &= " SET bank_name='" & lsBankName & "',"
                    lsSQL &= " bank_update_by=0,"
                    lsSQL &= " bank_update_date=SYSDATETIME(),"
                    lsSQL &= " HRIS_LASTMODIFIEDON='" & loMigrationData.Item("LAST_MODIFIED_ON") & "' "
                    lsSQL &= " WHERE LOWER(bank_code)='" & lsBankCode.ToLower & "'"

                    lbIsModified = True
                End If

                lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                If lsSQL = "" Then
                    If lbIsModified Then
                        liModifiedCount += 1
                    Else
                        liNewCount += 1
                    End If
                Else
                    liErrored += 1
                End If
            End While

            SummaryInsert("BANK", liNewCount, liModifiedCount, liErrored)

        End If
    End Sub

    Private Function SummaryInsert(ByVal psMode As String, ByVal piNew As Integer, ByVal piModified As Integer, ByVal piErrored As Integer)
        Dim lsSQL As String
        lsSQL = ""
        lsSQL &= " INSERT INTO iem_mig_tflids(FLIDS_DATE, FLIDS_UPDATEAT, FLIDS_NEWINSERT, FLIDS_MODIFIED, FLIDS_ERRORED) "
        lsSQL &= " VALUES(SYSDATETIME(),'" & psMode & "'," & piNew & "," & piModified & "," & piErrored & ")"

        lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

    End Function
    Private Function IsExistsAtIEM(ByVal psDestination As String, ByVal psCode As String) As Boolean
        Dim lsSQl As String

        IsExistsAtIEM = False

        If psDestination = "BANK" Then
            lsSQl = ""
            lsSQl &= " SELECT bank_code "
            lsSQl &= " FROM iem_mst_tbank "
            lsSQl &= " WHERE bank_code='" & psCode & "' "

            IsExistsAtIEM = IIf(loDBConnection.GetExecuteScalar(lsSQl).ToString = "", False, True)

        ElseIf psDestination = "FCCC" Then
            lsSQl = ""
            lsSQl &= " SELECT fccc_code "
            lsSQl &= " FROM iem_mst_tfccc "
            lsSQl &= " WHERE fccc_code='" & psCode & "' "

            IsExistsAtIEM = IIf(loDBConnection.GetExecuteScalar(lsSQl).ToString = "", False, True)


        ElseIf psDestination = "OU" Then
            lsSQl = ""
            lsSQl &= " SELECT ou_code "
            lsSQl &= " FROM iem_mst_tou "
            lsSQl &= " WHERE ou_code='" & psCode & "' "

            IsExistsAtIEM = IIf(loDBConnection.GetExecuteScalar(lsSQl).ToString = "", False, True)

        ElseIf psDestination = "PRODUCT" Then
            lsSQl = ""
            lsSQl &= " SELECT product_code "
            lsSQl &= " FROM iem_mst_tproduct "
            lsSQl &= " WHERE product_code='" & psCode & "' "

            IsExistsAtIEM = IIf(loDBConnection.GetExecuteScalar(lsSQl).ToString = "", False, True)

        ElseIf psDestination = "REGION" Then
            lsSQl = ""
            lsSQl &= " SELECT region_name "
            lsSQl &= " FROM iem_mst_tregion "
            lsSQl &= " WHERE LOWER(region_name)='" & psCode.ToLower & "' "

            IsExistsAtIEM = IIf(loDBConnection.GetExecuteScalar(lsSQl).ToString = "", False, True)

        ElseIf psDestination = "COUNTRY" Then
            lsSQl = ""
            lsSQl &= " SELECT country_code"
            lsSQl &= " FROM iem_mst_tcountry "
            lsSQl &= " WHERE LOWER(country_code)='" & psCode.ToLower & "' "

            IsExistsAtIEM = IIf(loDBConnection.GetExecuteScalar(lsSQl).ToString = "", False, True)

        ElseIf psDestination = "GRADE" Then
            lsSQl = ""
            lsSQl &= " SELECT grade_code "
            lsSQl &= " FROM iem_mst_tgrade "
            lsSQl &= " WHERE LOWER(grade_code)='" & psCode.ToLower & "' "

            IsExistsAtIEM = IIf(loDBConnection.GetExecuteScalar(lsSQl).ToString = "", False, True)

        ElseIf psDestination = "DEPARTMENT" Then
            lsSQl = ""
            lsSQl &= " SELECT dept_code "
            lsSQl &= " FROM iem_mst_tdept "
            lsSQl &= " WHERE LOWER(dept_code)='" & psCode.ToLower & "' "

            IsExistsAtIEM = IIf(loDBConnection.GetExecuteScalar(lsSQl).ToString = "", False, True)


        ElseIf psDestination = "HOLIDAY" Then
            lsSQl = ""
            lsSQl &= " SELECT holiday_date"
            lsSQl &= " FROM iem_mst_tholiday "
            lsSQl &= " WHERE holiday_date='" & Format(CDate(psCode), "yyyy-MM-dd hh:mm:ss") & "' "

            IsExistsAtIEM = IIf(loDBConnection.GetExecuteScalar(lsSQl).ToString = "", False, True)

        ElseIf psDestination = "EMPLOYEE" Then
            lsSQl = ""
            lsSQl &= " SELECT employee_code "
            lsSQl &= " FROM iem_mst_temployee "
            lsSQl &= " WHERE employee_code='" & psCode & "' "

            IsExistsAtIEM = IIf(loDBConnection.GetExecuteScalar(lsSQl).ToString = "", False, True)
        End If

    End Function
End Class
