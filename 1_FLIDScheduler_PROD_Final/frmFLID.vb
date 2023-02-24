Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Net
Imports System.IO

Public Class frmFLID

    Dim loDBConnection As New iODBCconnection 

    Dim liRegionID As Integer
    Dim liCityID As Integer
    Dim liCountryID As Integer
    Dim liStateID As Integer
    Dim liFLIDSGID As Integer
    Dim isParticularEmployee As String
    Dim EmployeeCode As String
    Dim lsLinkedServer As String
    Dim Count As Integer
    Dim FICC_BS_SEGMENT As String
    Dim HFC_BS_SEGMENT As String
    Dim BS_SEGMENT As String
    Dim HFC_FICC_BS_SEGMENT As String
    Dim BranchGID_BranchCode_ProductGID_ProductCode_FCGID_FCCode As String
    Dim Client As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim e2 As New ErrorLogger()
        Try
            'WriteLog("EXE Started")
            HFC_FICC_BS_SEGMENT = ConfigurationManager.AppSettings("HFC_FICC_BS_SEGMENT")
            'FICC_BS_SEGMENT = "'Central','Digital','Rural','Urban'"
            'HFC_BS_SEGMENT = "'Central','Central - HFC','Rural - HFC','Urban - HFC','Digital','Rural','Urban'"
            'WriteLog("EXE Started")
            'Count = 1
            'For Count As Integer = 1 To 3
            Client = ConfigurationManager.AppSettings("Client")
            ' If (Client = "HFC") Then
            'loDBConnection.HFCOpenConnection("", "", "", "") 'FIC
            'BS_SEGMENT = ConfigurationManager.AppSettings("BS_SEGMENT")
            'Else
            loDBConnection.OpenConnection("", "", "", "")
            BS_SEGMENT = ConfigurationManager.AppSettings("BS_SEGMENT")
            'End If

            'HFCDBConnection.OpenConnection("", "", "", "") 'HFC
            lsLinkedServer = ConfigurationManager.AppSettings("Integration")
            isParticularEmployee = ConfigurationManager.AppSettings("ParticularEmployee")
            EmployeeCode = ConfigurationManager.AppSettings("empcode")

            'EmployeeCode = "'180267','180320','180347','180636','180645','180796','180803','180946','180947','180998','P116858','P116859','P116860','P116861','P116862','P116982'"

            'WriteLog("Read App Config")
            Dim lsSQL As String
            lsSQL = ""
            lsSQL &= " INSERT INTO iem_trn_tflids(flids_date, flids_description) "
            lsSQL &= " VALUES(SYSDATETIME(),'E-Connect Data Integration')"

            lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

            lsSQL = ""
            lsSQL &= " SELECT MAX(flids_gid) FROM iem_trn_tflids "

            liFLIDSGID = loDBConnection.GetExecuteScalar(lsSQL)


            Me.Text = "BANK "
            Application.DoEvents()
            FLIDS_BANK()
            'e2.WriteToErrorLog("Logging", "logging", "bank")
            Me.Text = "GRADE"
            Application.DoEvents()

            FLIDS_GRADE()
            Me.Text = "EMPLOYEE"
            Application.DoEvents()
            Try
                FLIDS_EMPLOYEE()
            Catch ex As Exception

            End Try

            Try
                If (Client = "HFC") Then
                    UpdateCOAValuesforHFC(HFC_FICC_BS_SEGMENT)
                End If
            Catch ex As Exception

            End Try
            'e2.WriteToErrorLog("Logging", "logging", "employee")
            lsSQL = ""
            lsSQL = " UPDATE iem_trn_tflids SET flids_endedon=SYSDATETIME() WHERE flids_gid = " & liFLIDSGID

            lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)
            MsgBox("Success")
            '    Count = Count + 1
            'Next
            UpdateIEMIssues()
            End

        Catch ex As Exception
            WriteLog(ex.Message + " test - " + ex.InnerException.ToString())
            Dim el As New ErrorLogger()
            Dim ssll As String = ConfigurationManager.AppSettings("ssl")
            Dim send_mail As String = ConfigurationManager.AppSettings("sendmail")
            Dim send_pw As String = ConfigurationManager.AppSettings("sendpw")
            Dim smpt_Mail As String = ConfigurationManager.AppSettings("smptMail")
            Dim smpt_Mailport As String = ConfigurationManager.AppSettings("smptMailPort")
            Dim FromAddress As String = ConfigurationManager.AppSettings("From")
            Dim ToAddress As String = ConfigurationManager.AppSettings("To")

            Dim SmtpServer As New SmtpClient()
            Dim mail As New MailMessage()
            SmtpServer.Credentials = New  _
            Net.NetworkCredential(send_mail, send_pw)
            SmtpServer.Port = smpt_Mailport
            SmtpServer.Host = send_mail
            mail = New MailMessage()
            mail.From = New MailAddress(FromAddress)
            mail.To.Add(ToAddress)
            mail.Subject = "FLID"
            mail.Body = "Flids Problem Occured" + ex.Message.ToString
            SmtpServer.Send(mail)
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")

        End Try

    End Sub

    Private Sub FLIDS_EMPLOYEE()
        Try
            'WriteLog("test")

            Dim loMigrationData As SqlDataReader
            Dim loBankData As SqlDataReader
            Dim loGetData As SqlDataReader

            'Dim e3 As New ErrorLogger()
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
            Dim OFFICE_MAIL_ID As String
            Dim PERSONAL_EMAIL_ID As String
            Dim CONTACT_NO As String
            Dim MOBILE_NO As String
            Dim EMP_ACCOUNTNO As String
            Dim EMP_BANKNAME As String
            Dim EMP_IFS_NO As String
            Dim BRANCH_CODE As String
            Dim EMP_REPORTINGTO As String
            Dim DATE_OF_RESIGNATION As String
            Dim EMPLOYEE_STATUS As String
            Dim PRODUCT As String
            Dim lsLocation As String
            Dim EFFECTIVEDATE As String
            Dim CC, CCNAME As String
            Dim BS, BSNAME As String
            Dim PRODUCTDESCRIPTION As String
            Dim ACCOUNTTYPE As String
            Dim EMPLASTWORKINGDATE As String

            Dim liGradeGID As Integer
            Dim liDesignationGID As Integer
            Dim liDepartmentGID As Integer


            Dim liBankGID As Integer
            Dim liBSGID, liCCGID As Integer
            Dim liProductGID As Integer
            Dim liOUGID As Integer
            Dim liBranchGID As Integer
            Dim liEmployeeGID As Integer
            Dim lsFunctionName As String

            Dim lsOUCODE As String
            Dim lsQry As String
            Dim liErrored As Integer
            Dim lbIsModified As Boolean
            EMP_STAFFID = "0"
            Dim liPresentBranchGID As Integer
            'Ramya
            Dim lsLastModifiedOnForBank As String
            'Dim FirstTime As String

            'WriteLog("before getting last modified datetime")
            lsSQL = ""
            lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
            lsSQL &= " FROM iem_mst_temployee "

            lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL).ToString.Trim

            'lsLastModifiedOn = "2019-10-17"
            'WriteLog(lsSQL)

            lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)
            'WriteLog("last modified - " + lsLastModifiedOn.ToString)
            'Ramya
            lsLastModifiedOnForBank = lsLastModifiedOn
            lsTableName = lsLinkedServer & "[IEM_EMPLOYEE_FIELDS_FIC] "
            'FirstTime = ConfigurationManager.AppSettings("FirstTime")
            BS_SEGMENT = ConfigurationManager.AppSettings("BS_SEGMENT")
            'BS_SEGMENT = "Central"
            liNewCount = 0
            liModifiedCount = 0
            liErrored = 0
            'WriteLog(BS_SEGMENT)
            'EmployeeCode = "'180310','180312','180313','180314','180315','180316','180267','180318','180330','180262','180262','180167','180337','180339','180342','P116860','P116861','180346','P116862','180354'"
            lsSQL = ""
            lsSQL &= " SELECT * FROM " & lsTableName
            If (isParticularEmployee = "Y") Then
                lsSQL &= " WHERE emp_staffid in ('" & EmployeeCode & "')"
                'lsSQL &= " WHERE emp_staffid in ('P116859','180309','180308')"
            Else
                lsSQL &= " WHERE (LAST_MODIFIED_DATE IS NULL "
                lsSQL &= " OR LAST_MODIFIED_DATE >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' ) and BUSINESS_SEGMENT in (" & BS_SEGMENT & ")"
                'lsSQL &= " OR LAST_MODIFIED_DATE >='2021-10-01')  and BUSINESS_SEGMENT in (" & BS_SEGMENT & ")"
            End If



            'lsSQL &= " OR LAST_MODIFIED_DATE >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' ) and BUSINESS_SEGMENT in (" & BS_SEGMENT & ")"

            '151542','151716','152733','154736','156429','161695','163660','164165','164168'
            'lsSQL &= " where emp_staffid not in (select employee_code from iem_mst_temployee)"
            'lsSQL &= " and BUSINESS_SEGMENT in (" & BS_SEGMENT & ")"
            'lsSQL &= " WHERE emp_staffid in ('151542','151716','152733','154736','156429','161695','163660','164165','164168')"
            'lsSQL &= " where emp_staffid not in (select employee_code from iem_mst_temployee)"
            'lsSQL &= " and business_segment in ('central','central - hfc','urban - hfc','rural - hfc')"
            WriteLog(lsSQL)
            'If (isParticularEmployee = "N") & (FirstTime = "N") Then
            '    lsSQL &= " WHERE (LAST_MODIFIED_DATE IS NULL "
            '    lsSQL &= " OR LAST_MODIFIED_DATE >'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' ) and BUSINESS_SEGMENT in (" & BS_SEGMENT & ")"
            'ElseIf (isParticularEmployee = "N") & (FirstTime = "Y") Then
            '    lsSQL &= " WHERE BUSINESS_SEGMENT in (" & BS_SEGMENT & ")"
            'Else
            '    lsSQL &= " WHERE emp_staffid in ('" & EmployeeCode & "')"
            'End If

            loMigrationData = loDBConnection.GetDataReader(lsSQL)
            'WriteLog(lsSQL)
            If loMigrationData.HasRows Then

                Me.Text = "EMPLOYEE"

                While loMigrationData.Read

                    Try

                        EMP_STAFFID = loMigrationData.Item("EMP_STAFFID") & "".Trim                 ' [EMP_STAFFID]
                        'WriteLog(EMP_STAFFID)
                        EMPLASTWORKINGDATE = loMigrationData.Item("LAST_WORKING_DATE") & "".Trim    ' [LAST_WORKING_DATE]

                        lsLastModifiedOn = loMigrationData.Item("LAST_MODIFIED_DATE") & "".Trim     ' [LAST_MODIFIED_DATE]
                        'WriteLog(lsLastModifiedOn)
                        EFFECTIVEDATE = loMigrationData.Item("EFFECTIVE_DATE") & "".Trim            ' [EFFECTIVE DATE]

                        EMPLOYEE_STATUS = loMigrationData.Item("EMPLOYEE_STATUS") & "".Trim         ' [EMPLOYEE_STATUS]
                        EMPLOYEE_STATUS = IIf(EMPLOYEE_STATUS = "Active", "Y", "N")

                        EMP_COMPANYNAME = loMigrationData.Item("EMP_COMPANYNAME") & "".Trim         ' [EMP_COMPANYNAME]
                        EMP_COMPANYNAME = FormatTextInput(EMP_COMPANYNAME)
                        EMP_COMPANYNAME = Mid(EMP_COMPANYNAME, 1, 256)

                        DATE_OF_BIRTH = loMigrationData.Item("DATE_OF_BIRTH") & "".Trim             ' [DATE_OF_BIRTH]

                        PER_GENDER = loMigrationData.Item("PER_GENDER") & "".Trim                   ' [PER_GENDER]
                        PER_GENDER = Mid(PER_GENDER, 1, 1)

                        CON_HOME_ADDRESS1 = loMigrationData.Item("CON_HOME_ADDRESS1") & "".Trim     ' [CON_HOME_ADDRESS1]
                        CON_HOME_ADDRESS1 = FormatTextInput(CON_HOME_ADDRESS1, True)
                        CON_HOME_ADDRESS1 = Mid(CON_HOME_ADDRESS1, 1, 256)

                        CON_HOME_ADDRESS2 = loMigrationData.Item("CON_HOME_ADDRESS2") & "".Trim     ' [CON_HOME_ADDRESS2]
                        CON_HOME_ADDRESS2 = FormatTextInput(CON_HOME_ADDRESS2, True)
                        CON_HOME_ADDRESS2 = Mid(CON_HOME_ADDRESS2, 1, 256)

                        CON_HOME_ADDRESS3 = loMigrationData.Item("CON_HOME_ADDRESS3") & "".Trim     ' [CON_HOME_ADDRESS3]
                        CON_HOME_ADDRESS3 = FormatTextInput(CON_HOME_ADDRESS3, True)
                        CON_HOME_ADDRESS3 = Mid(CON_HOME_ADDRESS3, 1, 256)

                        CON_HOME_ADDRESS4 = loMigrationData.Item("CON_HOME_ADDRESS4") & "".Trim     ' [CON_HOME_ADDRESS4]
                        CON_HOME_ADDRESS4 = FormatTextInput(CON_HOME_ADDRESS4, True)
                        CON_HOME_ADDRESS4 = Mid(CON_HOME_ADDRESS4, 1, 256)

                        liRegionID = 0
                        liCountryID = 0
                        liStateID = 0
                        liCityID = 0

                        REGION_NAME = loMigrationData.Item("REGION_NAME") & "".Trim                 ' [REGION_NAME]
                        REGION_NAME = Mid(REGION_NAME, 1, 16)
                        liRegionID = GetRegionGID(REGION_NAME)
                        REGION_NAME = FormatTextInput(REGION_NAME)

                        COUNTRY_NAME = loMigrationData.Item("COUNTRY_NAME") & "".Trim               ' [COUNTRY_NAME] 
                        liCountryID = GetCountryGID(COUNTRY_NAME)
                        COUNTRY_NAME = FormatTextInput(COUNTRY_NAME)


                        STATE_NAME = loMigrationData.Item("STATE_NAME") & "".Trim                   ' [STATE_NAME] 
                        liStateID = GetStateGID(STATE_NAME, liRegionID, liCountryID)
                        STATE_NAME = FormatTextInput(STATE_NAME)

                        CITY_NAME = loMigrationData.Item("CITY_NAME") & "".Trim                     ' [CITY_NAME]
                        liCityID = GetCityGID(CITY_NAME, liRegionID, liCountryID, liStateID)
                        CITY_NAME = FormatTextInput(CITY_NAME)


                        PIN_CODE = loMigrationData.Item("PIN_CODE") & "".Trim                       ' [PIN_CODE]
                        PIN_CODE = Mid(PIN_CODE, 1, 8)

                        DATE_OF_JOINING = loMigrationData.Item("DATE_OF_JOINING") & "".Trim         ' [DATE_OF_JOINING]

                        GRADE_NAME = loMigrationData.Item("GRADE_NAME") & "".Trim                   ' [GRADE_NAME]
                        liGradeGID = GetGradeGID(GRADE_NAME)

                        'GRADE_NAME = Mid(GRADE_NAME.Trim, 1, 8)

                        GRADE_NAME = FormatTextInput(GRADE_NAME.Trim)  ' ramya modified on 20 may 22 to avoid letters cut after 8chars like Presiden

                        HRIS_DESIGNATION = loMigrationData.Item("HRIS_DESIGNATION") & "".Trim       ' [HRIS_DESIGNATION]
                        HRIS_DESIGNATION = FormatTextInput(HRIS_DESIGNATION)

                        IEM_DESIGNATION = loMigrationData.Item("IEM_DESIGNATION") & "".Trim         ' [IEM_DESIGNATION]
                        liDesignationGID = GetDesignationGID(IEM_DESIGNATION)

                        IEM_DESIGNATION = FormatTextInput(IEM_DESIGNATION)

                        DEPARTMENT = loMigrationData.Item("DEPARTMENT") & "".Trim                   ' [DEPARTMENT]
                        liDepartmentGID = GetDepartmentGID(DEPARTMENT)
                        DEPARTMENT = FormatTextInput(DEPARTMENT)

                        OFFICE_MAIL_ID = loMigrationData.Item("OFFICE_MAIL_ID") & "".Trim           ' [OFFICE_MAIL_ID]
                        OFFICE_MAIL_ID = FormatTextInput(OFFICE_MAIL_ID)

                        PERSONAL_EMAIL_ID = loMigrationData.Item("PERSONAL_EMAIL_ID") & "".Trim     ' [PERSONAL_EMAIL_ID]
                        PERSONAL_EMAIL_ID = FormatTextInput(PERSONAL_EMAIL_ID)

                        CONTACT_NO = loMigrationData.Item("CONTACT_NO") & "".Trim                   ' [CONTACT_NO]
                        MOBILE_NO = loMigrationData.Item("MOBILE_NO") & "".Trim                     ' [MOBILE_NO]


                        BRANCH_CODE = loMigrationData.Item("BRANCH_CODE") & "".Trim                 ' [BRANCH_CODE]
                        lsLocation = loMigrationData.Item("LOCATION") & "".Trim                     ' [LOCATION]

                        BS = loMigrationData.Item("BUSINESS_SEGMENT_CODE") & "".Trim                ' [BUSINESS_SEGMENT_CODE]
                        BSNAME = loMigrationData.Item("BUSINESS_SEGMENT") & "".Trim                 ' [BUSINESS_SEGMENT]

                        lsSQL = ""
                        lsSQL &= " SELECT fc_gid, fc_code FROM iem_mst_tfc "
                        lsSQL &= " WHERE fc_name = '" & BSNAME & "' "

                        loGetData = loDBConnection.GetDataReader(lsSQL)

                        If loGetData.HasRows Then
                            If loGetData.Read Then
                                liBSGID = loGetData.Item("fc_gid").ToString
                            End If
                        Else
                            liBSGID = 0
                        End If
                        Try
                            CC = loMigrationData.Item("CC_CODE") & "".Trim                              ' [CC_CODE]
                            CCNAME = loMigrationData.Item("CC_DESCRIPTION") & "".Trim                   ' [CC_DESCRIPTION]

                            If CCNAME.ToString.ToString <> "" Then
                                lsSQL = ""
                                lsSQL &= " SELECT cc_gid FROM iem_mst_tcc "
                                lsSQL &= " WHERE cc_code = '" & CC & "' "         'Mid(CCNAME, 1, 3)

                                liCCGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)
                            End If
                        Catch ex As Exception
                            WriteLog("CC Code error : " + lsSQL)
                        End Try

                        Try
                            PRODUCT = loMigrationData.Item("PRODUCT_CODE") & "".Trim                    ' [PRODUCT_CODE]
                            PRODUCTDESCRIPTION = loMigrationData.Item("PRODUCT_DESCRIPTION") & "".Trim  ' [PRODUCT_DESCRIPTION]

                            If PRODUCTDESCRIPTION.ToString.ToString <> "" Then
                                lsSQL = ""
                                lsSQL &= " SELECT product_gid FROM iem_mst_tproduct "
                                lsSQL &= " WHERE product_code = '" & PRODUCT & "' "                     ' Mid(PRODUCTDESCRIPTION, 1, 3)

                                liProductGID = Val(loDBConnection.GetExecuteScalar(lsSQL).ToString)
                            End If
                        Catch ex As Exception
                            WriteLog("Product Code error : " + lsSQL)
                        End Try
                        EMP_REPORTINGTO = loMigrationData.Item("EMP_REPORTINGTO") & "".Trim         ' [EMP_REPORTINGTO]
                        DATE_OF_RESIGNATION = loMigrationData.Item("DATE_OF_RESIGNATION") & "".Trim ' [DATE_OF_RESIGNATION]

                        liBranchGID = GetBranchID(BRANCH_CODE, "")
                        Try
                            If lsLocation <> "" Then
                                lsQry = " SELECT branch_code FROM iem_mst_tbranch "
                                lsQry &= " WHERE RIGHT(('0000' + branch_legacy_code),4) = '" & Mid(lsLocation.Trim, 1, 4) & "' "

                                lsOUCODE = loDBConnection.GetExecuteScalar(lsQry).ToString

                                liOUGID = GetBranchID(lsOUCODE, "")

                                liPresentBranchGID = liOUGID
                            Else
                                lsOUCODE = BRANCH_CODE
                                liOUGID = liBranchGID

                                liPresentBranchGID = liOUGID
                            End If
                        Catch ex As Exception
                            WriteLog("branch Code error : " + lsSQL)
                        End Try
                        lsOUCODE = BRANCH_CODE
                        liOUGID = liBranchGID

                        liEmployeeGID = GetEmployeeGID(EMP_REPORTINGTO)


                        'If IsDate(DATE_OF_RESIGNATION) Then
                        '    EMPLOYEE_STATUS = "R"
                        'End If

                        'If IsDate(EMPLASTWORKINGDATE) Then
                        '    If EMPLASTWORKINGDATE < Now Then EMPLOYEE_STATUS = "N"
                        'End If 

                        ''EMP_IFS_NO = "" 

                        Try
                            lsSQL = ""
                            lsSQL &= " SELECT B.Account_Type,A.* FROM " & lsLinkedServer & "[erm_employee_bank_details] A inner join " & lsLinkedServer & "[IEM_EMPLOYEE_BANK_DETAILS_FIC] B "
                            lsSQL &= " On A.Emp_StaffID=B.EMP_StaffID "
                            lsSQL &= " WHERE A.EMP_STAFFID='" & EMP_STAFFID.Trim & "' "


                            loBankData = loDBConnection.GetDataReader(lsSQL)

                            ACCOUNTTYPE = ""

                            If loBankData.HasRows Then

                                While loBankData.Read
                                    EMP_ACCOUNTNO = loBankData.Item("EMP_ACCOUNTNO") & "".Trim           ' [EMP_ACCOUNTNO]
                                    EMP_BANKNAME = loBankData.Item("EMP_BANKNAME") & "".Trim             ' [EMP_BANKNAME]

                                    'WriteLog("EMP_BANKNAME : " + Mid(EMP_BANKNAME, 1, 8))
                                    EMP_BANKNAME = Mid(EMP_BANKNAME, 1, 8)
                                    ACCOUNTTYPE = loBankData.Item("Account_Type") & "".Trim         ' [Account_Type]
                                    If (String.IsNullOrEmpty(loBankData.Item("EMP_IFS_NO").ToString())) Then
                                        EMP_IFS_NO = ""
                                    Else
                                        EMP_IFS_NO = loBankData.Item("EMP_IFS_NO") & "".Trim
                                    End If

                                    If ACCOUNTTYPE.Trim <> "" Then
                                        ACCOUNTTYPE = Mid(ACCOUNTTYPE, 1, 1)
                                    Else
                                        ACCOUNTTYPE = "S"
                                    End If
                                    'WriteLog("liBankGID : " + GetBankGID(EMP_BANKNAME).ToString())
                                    liBankGID = GetBankGID(EMP_BANKNAME)
                                    If ACCOUNTTYPE = "R" Then Exit While

                                End While
                            Else
                                ACCOUNTTYPE = ""
                                EMP_BANKNAME = ""
                                liBankGID = 0
                                EMP_ACCOUNTNO = ""
                                EMP_IFS_NO = ""
                            End If
                        Catch ex As Exception
                            WriteLog("branch Code error : " + lsSQL)
                            WriteLog("branch Code error : " + ex.Message)

                        End Try
                        lsFunctionName = ""

                        lsFunctionName = loMigrationData.Item("FUNCTION") & "".Trim            ' [EFFECTIVE DATE]

                        lsFunctionName = Mid(lsFunctionName, 1, 64)
                        'WriteLog("EMPStaffID - " + EMP_STAFFID)

                    Catch Ex As Exception
                        WriteLog("Employee Insert/Update loop error : " + Ex.Message + "EMPStaffID - " + EMP_STAFFID)
                        liErrored += 1

                        lsSQL = ""
                        lsSQL &= " INSERT INTO econ2iem_mst_temployee("
                        lsSQL &= " EMP_STAFFID, EMP_COMPANYNAME, DATE_OF_BIRTH, PER_GENDER, CON_HOME_ADDRESS1, CON_HOME_ADDRESS2, CON_HOME_ADDRESS3, CON_HOME_ADDRESS4"
                        lsSQL &= " CITY_NAME, STATE_NAME, REGION_NAME, COUNTRY_NAME, PIN_CODE, DATE_OF_JOINING, GRADE_NAME, HRIS_DESIGNATION, IEM_DESIGNATION,"
                        lsSQL &= " FUNCTION, OFFICE_MAIL_ID, PERSONAL_EMAIL_ID, CONTACT_NO, MOBILE_NO, BRANCH_CODE, Location, BUSINESS_SEGMENT_CODE, BUSINESS_SEGMENT, "
                        lsSQL &= " CC_CODE, CC_DESCRIPTION, BSCC_NAME, EMP_REPORTINGTO, DATE_OF_RESIGNATION, EMPLOYEE_STATUS, PRODUCT_CODE, PRODUCT_DESCRIPTION,"
                        lsSQL &= " OU_NAME, LAST_WORKING_DATE, LAST_MODIFIED_DATE, EFFECTIVE_DATE, DEPARTMENT, MSTATUS, ERRTEXT, flids_gid) VALUES("

                        lsSQL &= "'" & EMP_STAFFID & "','" & EMP_COMPANYNAME & "','" & DATE_OF_BIRTH & "','" & PER_GENDER & "','" & CON_HOME_ADDRESS1 & "',"
                        lsSQL &= "'" & CON_HOME_ADDRESS2 & "','" & CON_HOME_ADDRESS3 & "','" & CON_HOME_ADDRESS3 & "','" & CON_HOME_ADDRESS4 & "',"
                        lsSQL &= "'" & CITY_NAME & "','" & STATE_NAME & "','" & COUNTRY_NAME & "','" & PIN_CODE & "','" & DATE_OF_JOINING & "',"
                        lsSQL &= "'" & GRADE_NAME & "','" & HRIS_DESIGNATION & "','" & IEM_DESIGNATION & "','" & lsFunctionName & "','" & OFFICE_MAIL_ID & "',"
                        lsSQL &= "'" & PERSONAL_EMAIL_ID & "','" & CONTACT_NO & "','" & MOBILE_NO & "','" & BRANCH_CODE & "','" & lsLocation & "',"
                        lsSQL &= "'" & BS & "','" & BSNAME & "','" & CC & "','" & CCNAME & "','" & (BSNAME & "-" & CCNAME) & "',"
                        lsSQL &= "'" & EMP_REPORTINGTO & "','" & DATE_OF_RESIGNATION & "','" & EMPLOYEE_STATUS & "','" & PRODUCT & "','" & PRODUCTDESCRIPTION & "',"
                        lsSQL &= "'" & lsOUCODE & "','" & EMPLASTWORKINGDATE & "','" & lsLastModifiedOn & "','" & EFFECTIVEDATE & "',"
                        lsSQL &= "'" & DEPARTMENT & "','E','ERROR'," & liFLIDSGID & ")"

                        lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                        GoTo NextFetch
                    End Try

                    If GetMasterInfo("EMPLOYEE", EMP_STAFFID).ToString.ToString = "" Then
                        lsSQL = ""
                        lsSQL &= " INSERT INTO iem_mst_temployee(employee_code, employee_name, employee_dob, employee_gender, employee_addr1, "
                        lsSQL &= " employee_addr2, employee_addr3, employee_addr4, employee_city_name, employee_city_gid, "
                        lsSQL &= " employee_pincode, employee_doj, employee_grade_code, employee_grade_gid, employee_hris_designation, employee_iem_designation_gid, "
                        lsSQL &= " employee_dept_name, employee_dept_gid, employee_iem_designation,"   ', 
                        lsSQL &= " employee_office_email, employee_personal_email,"
                        lsSQL &= " employee_contact_no, employee_mobile_no, employee_era_acc_no, employee_era_bank_code, employee_era_bank_gid,"
                        lsSQL &= " employee_era_ifsc_code, employee_supervisor, "
                        lsSQL &= " employee_dor, employee_product_code, employee_product_gid, employee_ou_code, employee_ou_gid, "
                        lsSQL &= " employee_fc_code, employee_cc_code, employee_fccc_code, employee_effectivedate,  "
                        lsSQL &= " employee_branch_code, employee_branch_gid, employee_status,  employee_insert_by, employee_insert_date,employee_photo_flag,employee_photo_filename,employee_update_by, employee_update_date, "
                        lsSQL &= " HRIS_LASTMODIFIEDON, employee_unit_gid,employee_unit_name, employee_bankaccounttype, employee_functionname, "
                        lsSQL &= " employee_supervisor_code, employee_physical_branch_gid ) "  '
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
                        'lsSQL &= "'Migration',"

                        lsSQL &= "'" & CITY_NAME & "',"
                        lsSQL &= liCityID & ","
                        lsSQL &= "'" & PIN_CODE & "',"

                        If IsDate(DATE_OF_JOINING) Then
                            lsSQL &= "'" & Format(CDate(DATE_OF_JOINING), "yyyy-MM-dd") & "',"
                        Else
                            lsSQL &= "'1900-01-01',"
                        End If

                        lsSQL &= "'" & GRADE_NAME & "',"
                        lsSQL &= liGradeGID & ","

                        lsSQL &= "'" & HRIS_DESIGNATION & "'," & liDesignationGID & ","

                        lsSQL &= "'" & DEPARTMENT & "',"
                        lsSQL &= liDepartmentGID & ","

                        lsSQL &= "'" & IEM_DESIGNATION & "',"

                        lsSQL &= "'" & OFFICE_MAIL_ID & "',"
                        lsSQL &= "'" & PERSONAL_EMAIL_ID & "',"
                        lsSQL &= "'" & CONTACT_NO & "',"
                        lsSQL &= "'" & MOBILE_NO & "',"
                        lsSQL &= "'" & EMP_ACCOUNTNO & "',"

                        lsSQL &= "'" & EMP_BANKNAME & "',"
                        lsSQL &= liBankGID & ","

                        lsSQL &= "'" & EMP_IFS_NO & "',"
                        lsSQL &= liEmployeeGID & ","

                        If IsDate(DATE_OF_RESIGNATION) Then
                            lsSQL &= "'" & Format(CDate(DATE_OF_RESIGNATION), "yyyy-MM-dd") & "',"
                        Else
                            lsSQL &= "Null,"
                        End If

                        lsSQL &= "'" & PRODUCT & "',"
                        lsSQL &= liProductGID & ","

                        lsSQL &= "'" & BRANCH_CODE & "',"
                        lsSQL &= liBranchGID & ","

                        lsSQL &= "'" & BS & "','" & CC & "','" & BS & CC & "',"

                        If IsDate(EFFECTIVEDATE) Then
                            lsSQL &= "'" & Format(CDate(EFFECTIVEDATE), "yyyy-MM-dd") & "',"
                        Else
                            lsSQL &= "'1900-01-01',"
                        End If

                        lsSQL &= "'" & lsOUCODE & "',"
                        lsSQL &= liOUGID & ","


                        lsSQL &= "'" & EMPLOYEE_STATUS & "',"

                        lsSQL &= "0,SYSDATETIME(),'N','',0,SYSDATETIME(),"


                        'If IsDate(lsLastModifiedOn) Then
                        '    lsSQL &= "'" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "',"
                        'Else
                        '    lsSQL &= "'1900-01-01',"
                        'End If
                        'Ramya modified
                        Try
                            'lsSQL &= " HRIS_LASTMODIFIEDON='" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "',"
                            lsSQL &= "'" & lsLastModifiedOn & "',"
                            'WriteLog("Last before modified date format" + lsSQL)
                        Catch ex As Exception
                            lsSQL &= "'1900-01-01',"
                            WriteLog("Last modified date format" + ex.Message.ToString() + " - " + ex.InnerException.ToString())
                        End Try

                        'WriteLog(" before Insert Last modified date format" + lsSQL.ToString())
                        'lsSQL &= " HRIS_LASTMODIFIEDON='" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "',"
                        'WriteLog(" Insert after Last modified date format" + lsSQL.ToString())

                        lsSQL &= "0,'" & EMP_REPORTINGTO & "','" & ACCOUNTTYPE & "','" & lsFunctionName & "', "
                        lsSQL &= "'" & EMP_REPORTINGTO & "'," & liPresentBranchGID & ") "

                        lbIsModified = False

                    Else
                        lsSQL = ""
                        lsSQL &= " UPDATE iem_mst_temployee "

                        lsSQL &= " SET employee_name='" & EMP_COMPANYNAME & "',"

                        If IsDate(DATE_OF_BIRTH) Then
                            lsSQL &= "employee_dob='" & Format(CDate(DATE_OF_BIRTH), "yyyy-MM-dd") & "',"
                        Else
                            lsSQL &= "employee_dob='1900-01-01',"
                        End If

                        lsSQL &= " employee_gender='" & PER_GENDER & "', "
                        lsSQL &= " employee_addr1='" & CON_HOME_ADDRESS1 & "',"
                        lsSQL &= " employee_addr2='" & CON_HOME_ADDRESS2 & "',"
                        lsSQL &= " employee_addr3='" & CON_HOME_ADDRESS3 & "',"
                        'lsSQL &= " employee_addr4='" & CON_HOME_ADDRESS4 & "',"
                        lsSQL &= " employee_addr4='Migration',"
                        lsSQL &= " employee_city_name='" & CITY_NAME & "',"
                        lsSQL &= " employee_city_gid=" & liCityID & ","
                        lsSQL &= " employee_pincode='" & PIN_CODE & "',"

                        If IsDate(DATE_OF_JOINING) Then
                            lsSQL &= " employee_doj='" & Format(CDate(DATE_OF_JOINING), "yyyy-MM-dd") & "',"
                        Else
                            lsSQL &= " employee_doj='1900-01-01',"
                        End If
                        lsSQL &= " employee_grade_code='" & GRADE_NAME & "',"
                        lsSQL &= " employee_grade_gid=" & liGradeGID & ","
                        lsSQL &= " employee_hris_designation='" & HRIS_DESIGNATION & "',"
                        lsSQL &= " employee_iem_designation='" & IEM_DESIGNATION & "',"
                        lsSQL &= " employee_iem_designation_gid=" & liDesignationGID & ","
                        lsSQL &= " employee_dept_name='" & DEPARTMENT & "',"
                        lsSQL &= " employee_dept_gid=" & liDepartmentGID & ","
                        lsSQL &= " employee_unit_name='" & EMP_REPORTINGTO & "',"
                        lsSQL &= " employee_office_email='" & OFFICE_MAIL_ID & "',"
                        lsSQL &= " employee_personal_email='" & PERSONAL_EMAIL_ID & "',"
                        lsSQL &= " employee_contact_no='" & CONTACT_NO & "',"
                        lsSQL &= " employee_mobile_no='" & MOBILE_NO & "',"
                        lsSQL &= " employee_era_acc_no='" & EMP_ACCOUNTNO & "',"
                        lsSQL &= " employee_era_bank_code='" & EMP_BANKNAME & "',"
                        lsSQL &= " employee_era_bank_gid=" & liBankGID & ","
                        lsSQL &= " employee_era_ifsc_code='" & EMP_IFS_NO & "',"
                        lsSQL &= " employee_supervisor=" & liEmployeeGID & ","

                        If IsDate(DATE_OF_RESIGNATION) Then
                            lsSQL &= "employee_dor='" & Format(CDate(DATE_OF_RESIGNATION), "yyyy-MM-dd") & "',"
                        Else
                            lsSQL &= "employee_dor=Null,"
                        End If

                        lsSQL &= " employee_product_code='" & PRODUCT & "',"
                        lsSQL &= " employee_product_gid=" & liProductGID & ","
                        lsSQL &= " employee_ou_code='" & BRANCH_CODE & "',"
                        lsSQL &= " employee_ou_gid=" & liBranchGID & ","
                        lsSQL &= " employee_branch_code='" & lsOUCODE & "', "
                        lsSQL &= " employee_branch_gid=" & liOUGID & ","


                        lsSQL &= " employee_fc_code='" & BS & "',"
                        lsSQL &= " employee_cc_code='" & CC & "',"
                        lsSQL &= " employee_fccc_code='" & BS & CC & "',"

                        If IsDate(EFFECTIVEDATE) Then
                            'WriteLog(EFFECTIVEDATE)
                            lsSQL &= " employee_effectivedate='" & Format(CDate(EFFECTIVEDATE), "yyyy-MM-dd") & "',"
                        Else
                            lsSQL &= " employee_effectivedate=Null,"
                        End If


                        lsSQL &= " employee_status='" & EMPLOYEE_STATUS & "',"
                        lsSQL &= " employee_update_by=0, employee_isremoved='N',"
                        lsSQL &= " employee_update_date=SYSDATETIME(),"

                        'If IsDate(lsLastModifiedOn) Then
                        '    ' WriteLog(lsLastModifiedOn)

                        '    ' WriteLog("IF")
                        '    'WriteLog(Format(CDate(lsLastModifiedOn), "yyyy-MM-dd"))
                        '    lsSQL &= " HRIS_LASTMODIFIEDON='" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "',"
                        '    'WriteLog(lsSQL)
                        'Else
                        '    lsSQL &= " HRIS_LASTMODIFIEDON='1900-01-01',"
                        '    'WriteLog("else")
                        '    'WriteLog(lsSQL.ToString())
                        'End If
                        'Ramya modified
                        Try
                            'lsSQL &= " HRIS_LASTMODIFIEDON='" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "',"
                            lsSQL &= " HRIS_LASTMODIFIEDON='" & lsLastModifiedOn & "',"
                            'WriteLog("update Last modified date format" + lsSQL)
                        Catch ex As Exception
                            lsSQL &= " HRIS_LASTMODIFIEDON='1900-01-01',"
                            WriteLog("update Last modified date format" + ex.Message.ToString() + " - " + ex.InnerException.ToString())
                        End Try

                        'WriteLog("update before Last modified date format" + lsSQL)
                        'lsSQL &= " HRIS_LASTMODIFIEDON='" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "',"
                        'WriteLog("update after Last modified date format" + lsSQL)

                        lsSQL &= " employee_bankaccounttype='" & ACCOUNTTYPE & "',  "
                        lsSQL &= " employee_functionname='" & lsFunctionName & "', "

                        lsSQL &= " employee_supervisor_code='" & EMP_REPORTINGTO & "', "
                        lsSQL &= " employee_physical_branch_gid=" & liPresentBranchGID & ","

                        If IsDate(EMPLASTWORKINGDATE) Then
                            lsSQL &= " employee_lastworkingdate='" & Format(CDate(EMPLASTWORKINGDATE), "yyyy-MM-dd") & "' "
                        Else
                            lsSQL &= " employee_lastworkingdate=Null "
                        End If

                        lsSQL &= " WHERE ltrim(rtrim(employee_code))='" & EMP_STAFFID.Trim & "'"

                        lbIsModified = True
                        'WriteLog(lbIsModified.ToString())
                    End If
                    WriteLog("Employee Insert/Update query : " + lsSQL)
                    lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                    If lsSQL = "" Then
                        If lbIsModified Then
                            liModifiedCount += 1
                        Else
                            liNewCount += 1
                        End If
                    Else
                        liErrored += 1

                        Dim lsError As String = lsSQL
                        'WriteLog("Employee Insert/Update error : " + lsError)
                        '   lsSQL = ""
                        '   lsSQL &= " INSERT INTO econ2iem_mst_temployee("
                        '   lsSQL &= " EMP_STAFFID, EMP_COMPANYNAME, DATE_OF_BIRTH, PER_GENDER, CON_HOME_ADDRESS1, CON_HOME_ADDRESS2, CON_HOME_ADDRESS3, CON_HOME_ADDRESS4"
                        '   lsSQL &= " CITY_NAME, STATE_NAME, REGION_NAME, COUNTRY_NAME, PIN_CODE, DATE_OF_JOINING, GRADE_NAME, HRIS_DESIGNATION, IEM_DESIGNATION,"
                        '   lsSQL &= " FUNCTION, OFFICE_MAIL_ID, PERSONAL_EMAIL_ID, CONTACT_NO, MOBILE_NO, BRANCH_CODE, Location, BUSINESS_SEGMENT_CODE, BUSINESS_SEGMENT, "
                        '   lsSQL &= " CC_CODE, CC_DESCRIPTION, BSCC_NAME, EMP_REPORTINGTO, DATE_OF_RESIGNATION, EMPLOYEE_STATUS, PRODUCT_CODE, PRODUCT_DESCRIPTION,"
                        '   lsSQL &= " OU_NAME, LAST_WORKING_DATE, LAST_MODIFIED_DATE, EFFECTIVE_DATE, DEPARTMENT, MSTATUS, ERRTEXT, flids_gid) VALUES("

                        '  lsSQL &= "'" & EMP_STAFFID & "','" & EMP_COMPANYNAME & "','" & DATE_OF_BIRTH & "','" & PER_GENDER & "','" & CON_HOME_ADDRESS1 & "',"
                        '   lsSQL &= "'" & CON_HOME_ADDRESS2 & "','" & CON_HOME_ADDRESS3 & "','" & CON_HOME_ADDRESS3 & "','" & CON_HOME_ADDRESS4 & "',"
                        '   lsSQL &= "'" & CITY_NAME & "','" & STATE_NAME & "','" & COUNTRY_NAME & "','" & PIN_CODE & "','" & DATE_OF_JOINING & "',"
                        '   lsSQL &= "'" & GRADE_NAME & "','" & HRIS_DESIGNATION & "','" & IEM_DESIGNATION & "','" & lsFunctionName & "','" & OFFICE_MAIL_ID & "',"
                        '   lsSQL &= "'" & PERSONAL_EMAIL_ID & "','" & CONTACT_NO & "','" & MOBILE_NO & "','" & BRANCH_CODE & "','" & lsLocation & "',"
                        '   lsSQL &= "'" & BS & "','" & BSNAME & "','" & CC & "','" & CCNAME & "','" & (BSNAME & "-" & CCNAME) & "',"
                        '   lsSQL &= "'" & EMP_REPORTINGTO & "','" & DATE_OF_RESIGNATION & "','" & EMPLOYEE_STATUS & "','" & PRODUCT & "','" & PRODUCTDESCRIPTION & "',"
                        '   lsSQL &= "'" & lsOUCODE & "','" & EMPLASTWORKINGDATE & "','" & lsLastModifiedOn & "','" & EFFECTIVEDATE & "',"
                        '   lsSQL &= "'" & DEPARTMENT & "','E','" & lsError & "'," & liFLIDSGID & ")"

                        '  lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                    End If
NextFetch:
                    Me.Text = "EMPLOYEE : NEW - " & liNewCount & ", MODIFIED - " & liModifiedCount & ", Errored - " & liErrored & " ( " & EMP_STAFFID.Trim & " - " & EMP_COMPANYNAME.Trim & " ) "
                    Application.DoEvents()

                End While

                'lsSQL = " UPDATE iem_mst_temployee "
                'lsSQL &= " SET employee_status='N' WHERE employee_lastworkingdate<SYSDATETIME() "
                'WriteLog("Employee Insert/Update query : " + lsSQL)
                'lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)
                'WriteLog("Employee Insert/Update query : " + lsSQL)
                lsSQL = " UPDATE iem_mst_temployee "
                lsSQL &= " SET employee_supervisor = SUP.employee_gid "
                lsSQL &= " FROM iem_mst_temployee SUP "
                lsSQL &= " WHERE ltrim(rtrim(employee_supervisor_code))=ltrim(rtrim(SUP.employee_code)) "
                WriteLog("Employee Insert/Update query : " + lsSQL)
                lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                SummaryInsert("EMPLOYEE", liNewCount, liModifiedCount, liErrored)
                'Else
                '    WriteLog("No Records in Employee details")
            End If

            UpdateBankDetails(lsLastModifiedOnForBank)


        Catch ex As Exception
            WriteLog("Line No - 698: " + ex.Message + " -  next - " + ex.InnerException.ToString())
        End Try

        Application.Exit()
    End Sub

    Private Sub WriteLog(ByVal msg As String)
        Dim el As New ErrorLogger()
        el.WriteToErrorLog(msg, "", "Error")
       
    End Sub

    Private Sub UpdateCOAValuesforHFC(ByVal HFC_FICC_BS_SEGMENT As String)
        Dim lsSQL As String = ""
        Try
            'Dim loMigrationData As SqlDataReader 
            BranchGID_BranchCode_ProductGID_ProductCode_FCGID_FCCode = ConfigurationManager.AppSettings("BranchGID_BranchCode_ProductGID_ProductCode_FCGID_FCCode")
            Dim COAValues(6) As String
            COAValues = BranchGID_BranchCode_ProductGID_ProductCode_FCGID_FCCode.Split("_")

            lsSQL = ""
            lsSQL &= " UPDATE a set "
            lsSQL &= " employee_employeebranch_gid=" + COAValues(0) + "',employee_branch_gid=" + COAValues(0) + "',employee_branch_code='" + COAValues(1) + "',"
            lsSQL &= " employee_prev_branch_code='" + COAValues(1) + "',"
            lsSQL &= " employee_fccc_gid='" + COAValues(4) + "'+ Convert(varchar(3),b.cc_gid),"
            lsSQL &= " employee_fccc_code='" + COAValues(5) + "+ b.cc_code, "
            lsSQL &= " employee_fc_code='" + COAValues(5) + ",employee_product_gid=" + COAValues(2) + ",employee_product_code='" + COAValues(3) + "',employee_ou_gid=" + COAValues(0) + "',"
            lsSQL &= " employee_ou_code='" + COAValues(1) + "',employee_physical_branch_gid=" + COAValues(0) + ",employee_Client='FICC'"
            lsSQL &= "' from iem_mst_temployee a left join iem_mst_tcc b on a.employee_cc_code=b.cc_code and b.cc_isremoved ='N'"
            lsSQL &= " WHERE employee_fc_code in ('" + HFC_FICC_BS_SEGMENT + "')"
            'lsSQL &= " WHERE employee_fc_code in ('55','51','52')"
            'WriteLog("UPDATE iem_mst_temployee set : " + lsSQL)
            lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)
            'liModifiedCount += 1

        Catch ex As Exception
            WriteLog("UpdateCOAValuesforHFC Details error Line No - 773: " + ex.Message + " query: " + lsSQL)
        End Try
    End Sub
    Private Sub UpdateBankDetails(ByVal lsLastModifiedOn As String)
        Dim lsSQL As String = ""
        Try
            Dim loMigrationData As SqlDataReader
            Dim loBankData As SqlDataReader
            Dim HRIS_LastModifiedOn As String

            Dim liNewCount, liModifiedCount, liErrored As Integer

            Dim EMP_STAFFID As String
            Dim EMP_ACCOUNTNO As String
            Dim EMP_BANKNAME As String
            Dim ACCOUNTTYPE As String
            Dim liBankGID As Integer

            liNewCount = 0
            liModifiedCount = 0
            liErrored = 0
            EMP_STAFFID = "0"

            lsSQL = ""
            lsSQL &= " SELECT Distinct EMP_STAFFID FROM " & lsLinkedServer & "[erm_employee_bank_details] "
            lsSQL &= " WHERE 1=1 "
            If (isParticularEmployee = "N") Then
                lsSQL &= " and EMP_ModifiedOn >='" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "
                'lsSQL &= " and EMP_ModifiedOn >='2021-10-01') "
            Else
                lsSQL &= " and emp_staffid in ('" & EmployeeCode & "')"
            End If

            'WriteLog("Bank details started")
            'WriteLog("Fetch with last modified time : " + lsSQL)
            loMigrationData = loDBConnection.GetDataReader(lsSQL)
            'WriteLog("Fetch with last modified time : " + lsSQL)
            ACCOUNTTYPE = ""

            If loMigrationData.HasRows Then
                While loMigrationData.Read
                    Try
                        EMP_STAFFID = loMigrationData.Item("EMP_STAFFID") & "".Trim
                        lsSQL = ""
                        lsSQL &= " SELECT B.Account_Type,A.* FROM " & lsLinkedServer & "[erm_employee_bank_details] A inner join " & lsLinkedServer & "[IEM_EMPLOYEE_BANK_DETAILS_FIC] B "
                        lsSQL &= " On A.Emp_StaffID=B.EMP_StaffID "
                        lsSQL &= " WHERE A.EMP_STAFFID='" & EMP_STAFFID.Trim & "' "
                        'WriteLog("Fetch bank details with emp gid " + lsSQL)
                        loBankData = loDBConnection.GetDataReader(lsSQL)
                    Catch ex As Exception
                        WriteLog("Select Bank Details error Line No - 768: " + ex.Message + " query: " + lsSQL + "EMP_STAFFID: " + EMP_STAFFID)
                    End Try
                    'WriteLog("Fetch bank details with emp gid " + lsSQL)
                    If loBankData.HasRows Then
                        Try
                            While loBankData.Read

                                EMP_ACCOUNTNO = loBankData.Item("EMP_ACCOUNTNO") & "".Trim             ' [EMP_ACCOUNTNO]
                                EMP_BANKNAME = loBankData.Item("EMP_BANKNAME") & "".Trim               ' [EMP_BANKNAME]
                                liBankGID = GetBankGID(EMP_BANKNAME)
                                EMP_BANKNAME = Mid(EMP_BANKNAME, 1, 8)
                                lsLastModifiedOn = loBankData.Item("EMP_ModifiedOn") & "".Trim         ' [EMP_ModifiedOn]
                                ACCOUNTTYPE = loBankData.Item("Account_Type") & "".Trim                ' [ACCOUNT_TYPE]

                                If ACCOUNTTYPE.Trim <> "" Then
                                    ACCOUNTTYPE = Mid(ACCOUNTTYPE, 1, 1)
                                Else
                                    ACCOUNTTYPE = "S"
                                End If

                                If ACCOUNTTYPE = "R" Then Exit While
                            End While

                            lsSQL = ""
                            lsSQL &= " UPDATE iem_mst_temployee set "
                            lsSQL &= " employee_era_acc_no='" & EMP_ACCOUNTNO & "',"
                            lsSQL &= " employee_era_bank_code='" & EMP_BANKNAME & "',"
                            lsSQL &= " employee_era_bank_gid=" & liBankGID & ","
                            lsSQL &= " employee_bankaccounttype='" & ACCOUNTTYPE & "',"
                            HRIS_LastModifiedOn = ""
                            If IsDate(lsLastModifiedOn) Then
                                HRIS_LastModifiedOn = Format(CDate(lsLastModifiedOn), "yyyy-MM-dd")
                            Else
                                HRIS_LastModifiedOn = "1900-01-01"
                            End If
                            lsSQL &= " HRIS_LASTMODIFIEDON='" & HRIS_LastModifiedOn & "'"
                            lsSQL &= " WHERE ltrim(rtrim(employee_code))='" & EMP_STAFFID.Trim & "'"
                            'WriteLog("UPDATE iem_mst_temployee set : " + lsSQL)
                            lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)
                            liModifiedCount += 1

                        Catch Ex As Exception
                            liErrored += 1
                            Dim lsError As String = lsSQL
                            WriteLog("error while Bank update : " + lsError + " - " + Ex.Message + " - EmpCode : " + EMP_STAFFID)
                        End Try
                    End If


                End While
                Try
                    SummaryInsert("BANK", liNewCount, liModifiedCount, liErrored)
                Catch ex As Exception
                    WriteLog("Bank Details summary insert Line No - 921: " + ex.Message + " query: " + lsSQL + " - EmpCode : " + EMP_STAFFID)
                End Try


            End If
        Catch ex As Exception
            WriteLog("Bank Details error Line No - 818: " + ex.Message + " query: " + lsSQL)
        End Try
    End Sub

    Private Function GetEmployeeGID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT employee_gid FROM iem_mst_temployee "
        lsSQl &= " WHERE LOWER(employee_code)='" & psState.ToString.ToLower.Trim & "' "

        GetEmployeeGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

    End Function

    Private Function GetBranchID(ByVal psState As String, ByVal psName As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT branch_gid FROM iem_mst_tbranch "
        lsSQl &= " WHERE LOWER(branch_code)='" & psState.ToString.ToLower & "' AND branch_isremoved='N'  "

        GetBranchID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

    End Function

    Private Function GetBankGID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT bank_gid FROM iem_mst_tbank "
        lsSQl &= " WHERE LOWER(bank_name) Like '" & psState.ToString.ToLower & "%' AND bank_isremoved='N' "
        'WriteLog(lsSQl)
        GetBankGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        'WriteLog("GetBankGID : " + GetBankGID.ToString())
    End Function

    Private Function GetBS(ByVal psCode As String, ByVal psName As String) As Integer

        psName = FormatTextInput(psName)

        Dim lsSQl As String

        lsSQl = ""
        lsSQl &= " SELECT fc_gid FROM iem_mst_tfc "
        lsSQl &= " WHERE LOWER(fc_code)='" & psCode.ToString.ToLower & "' AND fc_isremoved='N' "
        ' lsSQl &= " AND LOWER(fc_name)='" & psName.ToString.ToLower & "'   "

        GetBS = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

    End Function

    Private Function GetCC(ByVal psCode As String, ByVal psName As String) As Integer

        psName = FormatTextInput(psName)

        If psName = "NULL" Then
            GetCC = 0
            Exit Function
        End If

        Dim lsSQl As String

        lsSQl = ""
        lsSQl &= " SELECT cc_gid FROM iem_mst_tcc "
        lsSQl &= " WHERE LOWER(cc_code)='" & psCode.ToString.ToLower & "'  AND cc_isremoved='N'  "
        'lsSQl &= " AND LOWER(cc_name)='" & Mid(psName.ToString.ToLower, 1, 32) & "'  "

        GetCC = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

    End Function

    Private Function GetProduct(ByVal psCode As String, ByVal psName As String) As Integer

        psName = FormatTextInput(psName)

        Dim lsSQl As String

        lsSQl = ""
        lsSQl &= " SELECT product_gid FROM iem_mst_tproduct "
        lsSQl &= " WHERE LOWER(product_code)='" & psCode.ToString.ToLower & "' AND product_isremoved='N'  "
        'lsSQl &= " AND LOWER(product_name)='" & Mid(psName.ToString.ToLower, 1, 32) & "'  "

        GetProduct = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        If GetProduct = 0 Then

            'Dim lsFCCCode = GetCode("iem_mst_tfc", "fccc_name", psState, 16, "fccc_code")

            lsSQl = ""
            lsSQl &= " INSERT INTO iem_mst_tproduct(product_code,product_name, product_insert_by, product_insert_date, "
            lsSQl &= " HRIS_LASTMODIFIEDON) VALUES('" & psCode & "','" & Mid(psName.ToString.ToLower, 1, 32) & "',0,SYSDATETIME(),SYSDATETIME()) "

            lsSQl = loDBConnection.ExecuteNonQuerySQL(lsSQl)

            lsSQl = ""
            lsSQl &= " SELECT product_gid FROM iem_mst_tproduct "
            lsSQl &= " WHERE LOWER(product_code)='" & psCode.ToString.ToLower & "' AND product_isremoved='N'  "
            'lsSQl &= " AND LOWER(product_name)='" & Mid(psName.ToString.ToLower, 1, 32) & "'  "

            GetProduct = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)
        End If

    End Function

    Private Function GetProductGID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)
        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT product_gid FROM iem_mst_tproduct "
        lsSQl &= " WHERE LOWER(product_name)='" & psState.ToString.ToLower & "' "

        GetProductGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

    End Function

    Private Function GetBusinessID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT unit_gid FROM iem_mst_tunit "
        lsSQl &= " WHERE LOWER(unit_name)='" & psState.ToString.ToLower & "' AND unit_isremoved='N' "

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
            lsSQl &= " WHERE LOWER(unit_name)='" & psState.ToString.ToLower & "' AND unit_isremoved='N' "

            GetBusinessID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)
        End If
    End Function

    Private Function GetDesignationGID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT designation_gid FROM iem_mst_tdesignation "
        lsSQl &= " WHERE LOWER(designation_name)='" & psState.ToString.ToLower & "' AND designation_isremoved='N' "

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
            lsSQl &= " WHERE LOWER(designation_name)='" & psState.ToString.ToLower & "' AND designation_isremoved='N' "

            GetDesignationGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)
        End If
    End Function

    Private Function GetDepartmentGID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT dept_gid FROM iem_mst_tdept "
        lsSQl &= " WHERE LOWER(dept_name)='" & psState.ToString.ToLower & "' AND dept_isremoved='N' "

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
            lsSQl &= " WHERE LOWER(dept_name)='" & psState.ToString.ToLower & "' AND dept_isremoved='N' "

            GetDepartmentGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)
        End If
    End Function

    Private Function FormatTextInput(ByVal psInput As String, Optional ByVal IgnoreQuotes As Boolean = False) As String

        If IgnoreQuotes Then
            psInput = psInput.Replace("'", "")
            psInput = psInput.Replace("\", "")
        Else
            psInput = psInput.Replace("'", "''")
            psInput = psInput.Replace("\", "\\")
        End If

        FormatTextInput = psInput
    End Function

    Private Function GetGradeGID(ByVal psState As String) As Integer

        psState = FormatTextInput(psState)

        Dim lsSQl As String
        lsSQl = ""
        lsSQl &= " SELECT grade_gid FROM iem_mst_tgrade "
        lsSQl &= " WHERE LOWER(grade_name)='" & psState.ToString.ToLower & "' "

        GetGradeGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

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
        lsSQl &= " WHERE LOWER(city_name)='" & psState.ToString.ToLower & "' AND city_isremoved='N' "

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
            lsSQl &= " WHERE LOWER(city_name)='" & psState.ToString.ToLower & "' AND city_isremoved='N'  "

            GetCityGID = Val(loDBConnection.GetExecuteScalar(lsSQl).ToString)

        End If
    End Function

    Private Sub FLIDS_HOLIDAY()
        Dim loMigrationData As SqlDataReader

        Dim lsLastModifiedOn As String
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

        lsSQL = ""
        lsSQL &= " SELECT MAX(HRIS_LASTMODIFIEDON) as HRIS_LASTMODIFIEDON "
        lsSQL &= " FROM iem_mst_tholiday "

        lsLastModifiedOn = loDBConnection.GetExecuteScalar(lsSQL)

        lsLastModifiedOn = IIf(lsLastModifiedOn = "", "01-01-1900", lsLastModifiedOn)

        lsTableName = lsLinkedServer & "[IEM_HOLIDAY_MASTER_FIC] "
        liNewCount = 0
        liModifiedCount = 0
        liErrored = 0
        lsSQL = ""
        lsSQL &= " SELECT * FROM " & lsTableName
        lsSQL &= " WHERE MODIFIED_ON IS NULL "
        lsSQL &= " OR MODIFIED_ON >='" & Format(CDate(lsLastModifiedOn), "yyyy-MM-dd") & "' "
        'lsSQL &= " OR MODIFIED_ON >='2021-10-01' "

        loMigrationData = loDBConnection.GetDataReader(lsSQL)

        If loMigrationData.HasRows Then

            While loMigrationData.Read

                ldHolidayDate = loMigrationData.Item("HOLIDAY_DATE")
                lsHolidayDescription = loMigrationData.Item("HOLIDAY_DESCRIPTION").ToString
                lsHolidayState = loMigrationData.Item("CALENDAR_NAME").ToString

                'Dim lsHoliday As String

                'lsHoliday = GetMasterInfo("HOLIDAY", ldHolidayDate)

                lsSQL = ""
                lsSQL &= " SELECT holiday_gid"
                lsSQL &= " FROM iem_mst_tholiday "
                lsSQL &= " WHERE holiday_date='" & Format(CDate(ldHolidayDate), "yyyy-MM-dd") & "' "
                If Val(loDBConnection.GetExecuteScalar(lsSQL).ToString) = 0 Then
                    lsSQL = ""
                    lsSQL &= " INSERT INTO iem_mst_tholiday(holiday_date, holiday_name, holiday_national_flag, holiday_state_flag, holiday_cutoff_flag,"
                    lsSQL &= " holiday_insert_by, holiday_insert_date, HRIS_LASTMODIFIEDON ) "
                    lsSQL &= " VALUES('" & Format(CDate(ldHolidayDate), "yyyy-MM-dd") & "',"
                    lsSQL &= "'" & lsHolidayDescription & "','Y','Y','Y',0,SYSDATETIME(),SYSDATETIME()) "

                    lbIsModified = False

                Else
                    lsSQL = ""
                    lsSQL &= " UPDATE iem_mst_tholiday "
                    lsSQL &= " SET holiday_name='" & lsHolidayDescription & "',"
                    lsSQL &= " holiday_update_by=0,"
                    lsSQL &= " holiday_update_date=SYSDATETIME() "
                    lsSQL &= " HRIS_LASTMODIFIEDON='" & loMigrationData.Item("MODIFIED_ON") & "' "
                    lsSQL &= " WHERE holiday_date='" & Format(CDate(loMigrationData.Item("HOLIDAY_DATE")), "yyyy-MM-dd") & "' "


                    lbIsModified = True
                End If

                lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                HolidayStateInsert(Format(CDate(ldHolidayDate), "yyyy-MM-dd hh:mm:ss"), lsHolidayState)

                If lsSQL = "" Then
                    If lbIsModified Then
                        liModifiedCount += 1
                    Else
                        liNewCount += 1
                    End If
                Else
                    liErrored += 1

                    Dim lsError As String = lsSQL

                    lsSQL = ""
                    lsSQL &= " INSERT INTO econ2iem_mst_tholiday(holiday_date, holiday_description, calendar_name, MODIFIED_ON,"
                    lsSQL &= " MSTATUS, ERRTEXT, flids_gid) VALUES ('" & Format(CDate(ldHolidayDate), "yyyy-MM-dd") & "','" & lsHolidayDescription & "','" & lsHolidayState & "','"
                    lsSQL &= loMigrationData.Item("MODIFIED_ON") & "','E','" & lsError & "'," & liFLIDSGID & ")"

                    lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                End If
            End While

            SummaryInsert("HOLIDAY", liNewCount, liModifiedCount, liErrored)

        End If
    End Sub

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

        lsTableName = lsLinkedServer & "[IEM_DESIGNATION_FIELDS_FIC] "
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

                lsDesignationCode = GetMasterInfo("DESIGNATION", lsDesignationName)

                If Not lsDesignationCode.ToString.Trim = "" Then
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
                    lsSQL &= " WHERE ltrim(rtrim(LOWER(designation_code)))='" & lsDesignationCode.ToLower.Trim & "'"

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

                    Dim lsError As String = lsSQL

                    lsSQL = ""
                    lsSQL &= " INSERT INTO econ2iem_mst_tdesignation(DESIGNATION_NAME, DESIGNATION_DESCRIPTION, DESIGNATION_LEVEL, LAST_MODIFIED_ON,"
                    lsSQL &= " MSTATUS, ERRTEXT, flids_gid) VALUES ('" & lsDesignationCode & "','" & lsDesignationName & "','" & lsDesignationLevel & "','"
                    lsSQL &= loMigrationData.Item("LAST_MODIFIED_ON") & "','E','" & lsError & "'," & liFLIDSGID & ")"

                    lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)


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

        lsTableName = lsLinkedServer & "[IEM_DEPARTMENT_FIELDS_FIC] "
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
                'lsDeptName = Mid(lsDeptName, 1, 32)
                'lsDeptCode = Mid(lsDeptCode, 1, 8)


                lsDeptCode = GetMasterInfo("DEPARTMENT", lsDeptName)

                If lsDeptCode.ToString.Trim = "" Then
                    lsDeptCode = GetCode("iem_mst_tdept", "dept_name", lsDeptName, 8, "dept_code")

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
                    lsSQL &= " WHERE ltrim(rtrim(LOWER(dept_code)))='" & lsDeptCode.ToLower.Trim & "'"

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

        lsTableName = lsLinkedServer & "[IEM_GRADE_FIELDS_FIC] "
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

                'lsGradeCode = GetCode("iem_mst_tgrade", "grade_name", lsGradeCode, 8, "grade_code")

                lsGradeCode = GetMasterInfo("GRADE", lsGradename)

                lsGradeLevel = loMigrationData.Item("GRADE_HIERARCHY").ToString.Trim
                lsGradeLevel = FormatTextInput(lsGradeLevel)

                If lsGradeCode.ToString.Trim = "" Then
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
                    lsSQL &= " WHERE ltrim(rtrim(LOWER(grade_code)))='" & lsGradeCode.ToLower.Trim & "'"

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

                    Dim lsError As String = lsSQL

                    lsSQL = ""
                    lsSQL &= " INSERT INTO econ2iem_mst_tgrade(grade_name, grade_description, grade_hierarchy, LAST_MODIFIED_ON,"
                    lsSQL &= " MSTATUS, ERRTEXT, flids_gid) VALUES ('" & lsGradeCode & "','" & lsGradename & "','" & lsGradeLevel & "','"
                    lsSQL &= loMigrationData.Item("LAST_MODIFIED_ON") & "','E','" & lsError & "'," & liFLIDSGID & ")"

                    lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

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

        lsTableName = lsLinkedServer & "[IEM_COUNTRY_FIELDS_FIC] "
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

                'lsCountryCode = GetCode("iem_mst_tcountry", "country_name", lsCountryName, 8, "country_code")

                lsCountryCode = GetMasterInfo("COUNTRY", lsCountryCode)

                If lsCountryCode.ToString.Trim = "" Then
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
                    lsSQL &= " WHERE ltrim(rtrim(LOWER(country_code)))='" & lsCountryCode.ToLower.Trim & "'"

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

        lsTableName = lsLinkedServer & "[IEM_REGION_FIELDS_FIC] "
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

                lsRegion = GetMasterInfo("REGION", lsRegion)

                If lsRegion.ToString.Trim = "" Then
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

        lsTableName = lsLinkedServer & "[IEM_PRODUCT_FIELDS_FIC] "
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

                lsProduct = Mid(lsName.ToString.Trim, 1, 3)

                lsProduct = GetMasterInfo("PRODUCT", lsProduct)

                If lsProduct.ToString.Trim = "" Then
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
                    lsSQL &= " WHERE ltrim(rtrim(product_code))='" & lsProduct.Trim & "'"

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

        lsTableName = lsLinkedServer & "[IEM_BANK_FIELDS_FIC] "

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


                If GetMasterInfo("BANK", loMigrationData.Item("BANK_CODE")).Trim = "" Then
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
                    lsSQL &= " WHERE ltrim(rtrim(LOWER(bank_code)))='" & lsBankCode.ToLower.Trim & "'"

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

                    Dim lsError As String = lsSQL

                    lsSQL = ""
                    lsSQL &= " INSERT INTO econ2iem_mst_tbank(BANK_CODE, BANK_NAME, LAST_MODIFIED_ON,"
                    lsSQL &= " MSTATUS, ERRTEXT, flids_gid) VALUES ('" & lsBankCode & "','" & lsBankName & "','"
                    lsSQL &= loMigrationData.Item("LAST_MODIFIED_ON") & "','E','" & lsError & "'," & liFLIDSGID & ")"

                    lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

                End If

            End While

            SummaryInsert("BANK", liNewCount, liModifiedCount, liErrored)

        End If
    End Sub

    Private Function SummaryInsert(ByVal psMode As String, ByVal piNew As Integer, ByVal piModified As Integer, ByVal piErrored As Integer)
        Dim lsSQL As String
        lsSQL = ""
        lsSQL &= " INSERT INTO iem_mig_tflids(FLIDS_DATE, FLIDS_UPDATEAT, FLIDS_NEWINSERT, FLIDS_MODIFIED, FLIDS_ERRORED, trn_flids_gid) "
        lsSQL &= " VALUES(SYSDATETIME(),'" & psMode & "'," & piNew & "," & piModified & "," & piErrored & "," & liFLIDSGID & ")"

        lsSQL = loDBConnection.ExecuteNonQuerySQL(lsSQL)

    End Function

    Private Function GetMasterInfo(ByVal psDestination As String, ByVal psCode As String) As String
        Dim lsSQl As String

        'IsExistsAtIEM = False

        If psDestination = "BANK" Then
            lsSQl = ""
            lsSQl &= " SELECT bank_code "
            lsSQl &= " FROM iem_mst_tbank "
            lsSQl &= " WHERE bank_code='" & psCode & "' "

        ElseIf psDestination = "FCCC" Then
            lsSQl = ""
            lsSQl &= " SELECT fccc_code "
            lsSQl &= " FROM iem_mst_tfccc "
            lsSQl &= " WHERE fccc_code='" & psCode & "' "

        ElseIf psDestination = "OU" Then
            lsSQl = ""
            lsSQl &= " SELECT ou_code "
            lsSQl &= " FROM iem_mst_tou "
            lsSQl &= " WHERE ou_code='" & psCode & "' "

        ElseIf psDestination = "PRODUCT" Then
            lsSQl = ""
            lsSQl &= " SELECT product_code "
            lsSQl &= " FROM iem_mst_tproduct "
            lsSQl &= " WHERE product_code='" & psCode & "' "

        ElseIf psDestination = "REGION" Then
            lsSQl = ""
            lsSQl &= " SELECT region_name "
            lsSQl &= " FROM iem_mst_tregion "
            lsSQl &= " WHERE LOWER(region_name)='" & psCode.ToLower & "' "

        ElseIf psDestination = "COUNTRY" Then
            lsSQl = ""
            lsSQl &= " SELECT country_code"
            lsSQl &= " FROM iem_mst_tcountry "
            lsSQl &= " WHERE LOWER(country_code)='" & psCode.ToLower & "' "

        ElseIf psDestination = "GRADE" Then
            lsSQl = ""
            lsSQl &= " SELECT grade_code "
            lsSQl &= " FROM iem_mst_tgrade "
            lsSQl &= " WHERE LOWER(grade_name)='" & psCode.ToLower & "' "

        ElseIf psDestination = "DEPARTMENT" Then
            lsSQl = ""
            lsSQl &= " SELECT dept_code "
            lsSQl &= " FROM iem_mst_tdept "
            lsSQl &= " WHERE LOWER(dept_name)='" & psCode.ToLower & "' "

        ElseIf psDestination = "DESIGNATION" Then
            lsSQl = ""
            lsSQl &= " SELECT designation_name "
            lsSQl &= " FROM iem_mst_tdesignation "
            lsSQl &= " WHERE LOWER(designation_name)='" & psCode.ToLower & "' "

        ElseIf psDestination = "HOLIDAY" Then
            lsSQl = ""
            lsSQl &= " SELECT holiday_date"
            lsSQl &= " FROM iem_mst_tholiday "
            lsSQl &= " WHERE holiday_date='" & Format(CDate(psCode), "yyyy-MM-dd hh:mm:ss") & "' "

        ElseIf psDestination = "EMPLOYEE" Then
            lsSQl = ""
            lsSQl &= " SELECT employee_code "
            lsSQl &= " FROM iem_mst_temployee "
            lsSQl &= " WHERE employee_code='" & psCode & "' "
        End If

        GetMasterInfo = loDBConnection.GetExecuteScalar(lsSQl).ToString.Trim


    End Function

    Private Sub frmFLID_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

    End Sub

    Private Sub frmFLID_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button1_Click(sender, e)
    End Sub

    Private Sub UpdateIEMIssues()
        Dim lsSQL As String
        Try
            lsSQL = ""
            lsSQL = loDBConnection.ExecuteNonQuerySP("pr_iem_UpdateIEMIssues", "AmountMismatch")
        Catch ex As Exception

        End Try
        Try
            lsSQL = ""
            lsSQL = loDBConnection.ExecuteNonQuerySP("pr_iem_UpdateIEMIssues", "InactiveEmployee")
        Catch ex As Exception

        End Try
        Try
            lsSQL = ""
            lsSQL = loDBConnection.ExecuteNonQuerySP("pr_iem_UpdateIEMIssues", "GSTGLisEmpty")
        Catch ex As Exception

        End Try
        Try
            lsSQL = ""
            lsSQL = loDBConnection.ExecuteNonQuerySP("pr_iem_UpdateIEMIssues", "SupplierActivation")
        Catch ex As Exception

        End Try
        Try
            lsSQL = ""
            lsSQL = loDBConnection.ExecuteNonQuerySP("pr_iem_UpdateIEMIssues", "GSTNMismatch")
        Catch ex As Exception

        End Try
        Try
            lsSQL = ""
            lsSQL = loDBConnection.ExecuteNonQuerySP("pr_iem_UpdateIEMIssues", "CreditlineGLMismatch")
        Catch ex As Exception

        End Try
        Try
            lsSQL = ""
            lsSQL = loDBConnection.ExecuteNonQuerySP("pr_iem_UpdateIEMIssues", "DebitlineGLMismatch")
        Catch ex As Exception

        End Try
    End Sub


End Class

