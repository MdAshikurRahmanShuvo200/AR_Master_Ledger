Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class frmReportViewer

    ' Form1 থেকে পাওয়া মানগুলো রাখার জন্য Variables
    Public FromDate As Date
    Public ToDate As Date
    Public FromCustomerId As String
    Public ToCustomerId As String
    Public FromAccSet As String
    Public ToAccSet As String
    Public Property ShowDetails As Boolean

    Private Sub frmReportViewer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' 1. Create Report Document Instance
            Dim rptDoc As New ReportDocument()

            ' Get the absolute path of the report from the application's startup directory (bin/Debug or bin/Release)
            Dim reportPath As String = IO.Path.Combine(Application.StartupPath, "Revised Customer_Ledger_Template_v2.rpt")

            ' Load the Crystal Report file dynamically using the generated path
            rptDoc.Load(reportPath)
            ' 3. Load the Report
            rptDoc.Load(reportPath)

            ' 4. Pass Database Credentials (SQL Server Logins)
            Dim connectionInfo As New ConnectionInfo()
            connectionInfo.ServerName = "localhost"
            connectionInfo.DatabaseName = "BNBDAT"
            connectionInfo.UserID = "sa"
            connectionInfo.Password = "1122"

            ' Apply connection to all tables in the report
            Dim tables As Tables = rptDoc.Database.Tables
            For Each table As Table In tables
                Dim tableLogOnInfo As TableLogOnInfo = table.LogOnInfo
                tableLogOnInfo.ConnectionInfo = connectionInfo
                table.ApplyLogOnInfo(tableLogOnInfo)
            Next

            ' 5. Pass Parameters EXACTLY matching Crystal Report Parameter Names
            rptDoc.SetParameterValue("FromDate", FromDate)
            rptDoc.SetParameterValue("ToDate", ToDate)
            rptDoc.SetParameterValue("FromCustomerId", FromCustomerId)
            rptDoc.SetParameterValue("ToCustomerId", ToCustomerId)
            rptDoc.SetParameterValue("FromAccSet", FromAccSet)
            rptDoc.SetParameterValue("ToAccSet", ToAccSet)

            ' === এই লাইনটি বাদ পড়েছিল, যা এখন যোগ করা হয়েছে ===
            rptDoc.SetParameterValue("ShowDetails", Me.ShowDetails)

            ' 6. Assign to ReportViewer control ONLY AFTER setting parameters and logon info
            CrystalReportViewer1.ReportSource = rptDoc
            CrystalReportViewer1.RefreshReport()

            ' 7. Open Form in Maximized Mode
            Me.WindowState = FormWindowState.Maximized

        Catch ex As Exception
            MessageBox.Show("Error loading report: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class