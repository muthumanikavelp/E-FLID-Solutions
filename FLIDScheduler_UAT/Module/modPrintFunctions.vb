Imports System.Drawing.Printing
Imports System.IO

Public Class modPrintFunctions
    Private printFont As Font
    Private streamToPrint As StreamReader
    'Private Shared filePath As String
    Private PrintMode As PrinterSettings
    Private PageSet As PageSettings

    Public Sub New()
        'Printing(gs_Filepath)
    End Sub

    ' The PrintPage event is raised for each page to be printed.
    Private Sub pd_PrintPage1(ByVal sender As Object, ByVal ev As PrintPageEventArgs)
        Dim linesPerPage As Single = 0
        Dim yPos As Single = 0
        Dim count As Integer = 0
        Dim leftMargin As Single = ev.MarginBounds.Left
        Dim topMargin As Single = ev.MarginBounds.Top
        Dim line As String = Nothing

        ' Calculate the number of lines per page.
        linesPerPage = ev.MarginBounds.Height / printFont.GetHeight(ev.Graphics)

        ' Iterate over the file, printing each line.
        While count < linesPerPage
            line = streamToPrint.ReadLine()
            If line Is Nothing Then
                Exit While
            End If
            yPos = topMargin + count * printFont.GetHeight(ev.Graphics)
            ev.Graphics.DrawString(line, printFont, Brushes.Black, leftMargin, _
                yPos, New StringFormat())
            count += 1
        End While

        ' If more lines exist, print another page.
        If Not (line Is Nothing) Then
            ev.HasMorePages = True
        Else
            ev.HasMorePages = False
        End If
    End Sub

    ' The PrintPage event is raised for each page to be printed.
    Private Sub pd_PrintPage(ByVal sender As Object, ByVal ev As PrintPageEventArgs)
        Dim linesPerPage As Single = 0
        Dim yPos As Single = 0
        Dim count As Integer = 0
        Dim leftMargin As Single = ev.MarginBounds.Left
        Dim topMargin As Single = ev.MarginBounds.Top
        Dim line As String = Nothing

        ' Calculate the number of lines per page.
        linesPerPage = ev.MarginBounds.Height / printFont.GetHeight(ev.Graphics)

        ' Iterate over the file, printing each line.
        While Not line = Chr(12)
            line = streamToPrint.ReadLine()
            If line Is Nothing Then
                Exit While
            End If

            If InStr(line, "<B>", CompareMethod.Text) = 0 Then

                'printFont = New Font("Courier New", 10, FontStyle.Regular)
                printFont = New Font("Courier New", 10, FontStyle.Bold)
            Else
                line = Replace(line, "<B>", "")
                printFont = New Font("Courier New", 10, FontStyle.Bold)
            End If

            If Not line = Chr(12) Then
                yPos = topMargin + count * printFont.GetHeight(ev.Graphics)
                ev.Graphics.DrawString(line, printFont, Brushes.Black, leftMargin, _
                    yPos, New StringFormat())
                count += 1
            End If

        End While

        ' If more lines exist, print another page.
        If line = Chr(12) Then
            ev.HasMorePages = True
        Else
            ev.HasMorePages = False
        End If
    End Sub


    ' Print the file.
    Public Sub Printing(ByVal filepath As String)
        Dim land_val As Integer = 90
        Dim page_val As Boolean
        Try
            'filePath = "test.txt"
            streamToPrint = New StreamReader(filepath)
            Try
                printFont = New Font("Courier New", 10)
                'page_val = PageSet.Landscape
                'PageSet.Landscape = page_val
                'land_val = PrintMode.LandscapeAngle

                Dim pd As New PrintDocument()
                AddHandler pd.PrintPage, AddressOf pd_PrintPage
                'pd.DefaultPageSettings.Landscape = True
                pd.DefaultPageSettings.Landscape = False
                Dim margins As New Margins(40, 50, 50, 50)
                'Dim margins As New Margins(100, 100, 100, 100)
                pd.DefaultPageSettings.Margins = margins
                pd.PrinterSettings.Copies = 1
                ' Print the document.
                pd.Print()
            Finally
                streamToPrint.Close()
            End Try
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub 'Printing    

    '' This is the main entry point for the application.
    'Public Shared Sub Main()
    '    Dim args() As String = System.Environment.GetCommandLineArgs()
    '    Dim sampleName As String = args(0)
    '    If args.Length <> 1 Then
    '        Console.WriteLine("Usage: " & sampleName & " <file path>")
    '        Return
    '    End If
    '    gsPrintFilepath = args(0)
    'End Sub
    ' This is the main entry point for the application.
    Public Shared Sub PrintMain()
        Dim args() As String = System.Environment.GetCommandLineArgs()
        Dim sampleName As String = args(0)
        If args.Length <> 1 Then
            Console.WriteLine("Usage: " & sampleName & " <file path>")
            Return
        End If
        gsPrintFilepath = args(0)
    End Sub
End Class

