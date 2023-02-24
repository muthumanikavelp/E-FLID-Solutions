Imports FlexiLibrary
Imports System.IO
Imports System.Data.Odbc

Module Module1
    Public gobjConnection As New iODBCconnection    ' Connection Object 
    'Public gobjSecurity As New iSecurity
    Public gsUserId As String = ""
    Public gUserName As String = ""

    'Public Sub main()
    '    Try
    '        With gobjSecurity
    '            .LoginDbApplication = "S"
    '            .DbApplication = "S"
    '            .LoginCaption = "KIT"                      ' Login Caption       
    '            .LoginSoftCode = "KIT"                     ' Login Software Code.
    '            .LoginSoftVersion = "1.0.0"

    '            If Not File.Exists(Application.StartupPath & "\AppConfig.ini") Then
    '                MessageBox.Show("Configuration File is Missing", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '                End
    '            End If

    '            GetConnectionString()

    '            If .LoginDbApplication = "S" Then
    '                '.LoginDbConnectionString = "Driver={SQL Server};Server=FLEX-F5-19\SQLEXPRESS;Database=kit;Uid=sa;Pwd=gnsa;MARS_Connection=yes;"
    '                '.LoginDbConnectionString = "Driver={SQL Server Native Client 10.0};Server=FLEX-F5-19\SQLEXPRESS;Database=kit;Uid=sa;Pwd=gnsa;MARS_Connection=yes;"
    '                .LoginDbConnectionString = "Driver={SQL Server Native Client 10.0};Server=" & .LoginDBIP & ";Database=" & .LoginDBName & ";Uid=" & .LoginDBUserName & ";Pwd=" & .LoginDBPassword & ";MARS_Connection=yes;"
    '                gobjConnection.OpenConnection(.LoginDbConnectionString)
    '            End If

    '            .ShowLoginDialog()

    '            If Not .LoginState Then
    '                gobjSecurity.TerminateApplication()
    '                End
    '            Else
    '                If .DbApplication = "S" Then
    '                    '.DbConnectionString = "Driver={SQL Server};Server=FLEX-F5-19\SQLEXPRESS;Database=kit;Uid=sa;Pwd=gnsa;MARS_Connection=yes;"
    '                    '.DbConnectionString = "Driver={SQL Server Native Client 10.0};Server=FLEX-F5-19\SQLEXPRESS;Database=kit;Uid=sa;Pwd=gnsa;MARS_Connection=yes;"
    '                    .DbConnectionString = "Driver={SQL Server Native Client 10.0};Server=" & .LoginDBIP & ";Database=" & .LoginDBName & ";Uid=" & .LoginDBUserName & ";Pwd=" & .LoginDBPassword & ";MARS_Connection=yes;"
    '                    gobjConnection.OpenConnection(.DbConnectionString)
    '                End If
    '            End If

    '            If .LoginFromProcess <> "" Then
    '                .LoginUserCode = iRoutines.Decryption(.LoginFromProcess)
    '            End If

    '        End With


    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End
    '    End Try
    'End Sub

 

    Public Function ChkDate(ByVal dt As String) As Boolean
        ChkDate = False


        Select Case Val(Left(dt, 2))
            Case Is > 31, Is < 1
                MessageBox.Show("Invalid Day !", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Function
        End Select

        Select Case Val(Mid(dt, 4, 2))
            Case Is > 12, Is < 1
                MessageBox.Show("Invalid Month !", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Function
        End Select

        Select Case ""
            ' Case Trim(Mid(dt, Len(dt) - 1, 4))
            Case Trim(Mid(dt, 7, 4))
                MessageBox.Show("Invalid Year !", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Function
        End Select

        'If Not IsDate(dt) Then
        '    MessageBox.Show("Invalid Date !", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        '    Exit Function
        'End If

        ChkDate = True
    End Function

    Public Function ConvDate(ByVal dt As String) As String
        ConvDate = Format(DateSerial(Val(Mid(dt, 7, 4)), Val(Mid(dt, 4, 2)), Val(Left(dt, 2))), "yyyy-MM-dd")
    End Function

End Module
