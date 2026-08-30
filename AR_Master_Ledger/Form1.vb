Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Drawing

Public Class Form1

    ' Connection string retrieved from App.config using the name "SageDBConn"
    Private ReadOnly connString As String = ConfigurationManager.ConnectionStrings("SageDBConn").ConnectionString

    ' Executed when Form1 loads
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Load ComboBox data from SQL View on startup
        LoadCustomerIDsFromView()
        LoadAccountSetsFromView()

        ' SET DEFAULT DATES ON FORM LOAD
        SetDefaultDates()
    End Sub

    ' Method to fetch the oldest transaction date and set default dates
    Private Sub SetDefaultDates()
        Try
            Using conn As New SqlConnection(connString)
                ' SQL query to find the absolute oldest Transaction Date from the view
                Dim query As String = "SELECT MIN([TransactionDate]) FROM vw_AR_TRANSACTION WHERE [TransactionDate] IS NOT NULL"
                Dim cmd As New SqlCommand(query, conn)

                conn.Open()
                Dim val As Object = cmd.ExecuteScalar()

                ' If a valid date is found, assign it to From DatePicker, otherwise fallback to 1990/01/01
                If val IsNot DBNull.Value AndAlso val IsNot Nothing Then
                    dtpFromDate.Value = Convert.ToDateTime(val).Date
                Else
                    dtpFromDate.Value = New Date(1990, 1, 1)
                End If
            End Using
        Catch ex As Exception
            ' Fallback date in case database connection fails
            dtpFromDate.Value = New Date(1990, 1, 1)
        End Try

        ' Set To DatePicker to today's current date by default
        dtpToDate.Value = DateTime.Today
    End Sub

    ' Method to fetch unique Customer IDs from the SQL View in ascending order (A-Z)
    Private Sub LoadCustomerIDsFromView()
        Try
            Using conn As New SqlConnection(connString)
                Dim query As String = "SELECT DISTINCT [Customer ID] FROM vw_AR_TRANSACTION WHERE [Customer ID] IS NOT NULL AND [Customer ID] <> '' ORDER BY [Customer ID] ASC"
                Dim cmd As New SqlCommand(query, conn)

                conn.Open()
                Dim reader As SqlDataReader = cmd.ExecuteReader()

                cmbFromCustomerID.Items.Clear()
                cmbToCustomerID.Items.Clear()

                While reader.Read()
                    Dim custId As String = reader("Customer ID").ToString().Trim()
                    cmbFromCustomerID.Items.Add(custId)
                    cmbToCustomerID.Items.Add(custId)
                End While

                reader.Close()
            End Using
        Catch ex As Exception
            MessageBox.Show("Failed to load Customer IDs: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Method to fetch unique Account Set IDs from the SQL View in ascending order (A-Z)
    Private Sub LoadAccountSetsFromView()
        Try
            Using conn As New SqlConnection(connString)
                Dim query As String = "SELECT DISTINCT [Account Set ID] FROM vw_AR_TRANSACTION WHERE [Account Set ID] IS NOT NULL AND [Account Set ID] <> '' ORDER BY [Account Set ID] ASC"
                Dim cmd As New SqlCommand(query, conn)

                conn.Open()
                Dim reader As SqlDataReader = cmd.ExecuteReader()

                cmbFromAccSetID.Items.Clear()
                cmbToAccSetID.Items.Clear()

                While reader.Read()
                    Dim accSetId As String = reader("Account Set ID").ToString().Trim()
                    cmbFromAccSetID.Items.Add(accSetId)
                    cmbToAccSetID.Items.Add(accSetId)
                End While

                reader.Close()
            End Using
        Catch ex As Exception
            MessageBox.Show("Failed to load Account Set IDs: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Helper function to fetch Min or Max value dynamically from database when inputs are empty
    Private Function GetFirstOrLastValue(ByVal columnName As String, ByVal isMin As Boolean) As String
        Dim result As String = ""
        Dim aggregate As String = If(isMin, "MIN", "MAX")
        Dim query As String = $"SELECT {aggregate}([{columnName}]) FROM vw_AR_TRANSACTION WHERE [{columnName}] IS NOT NULL AND [{columnName}] <> ''"

        Try
            Using conn As New SqlConnection(connString)
                Dim cmd As New SqlCommand(query, conn)
                conn.Open()
                Dim val As Object = cmd.ExecuteScalar()
                If val IsNot DBNull.Value AndAlso val IsNot Nothing Then
                    result = val.ToString().Trim()
                End If
            End Using
        Catch ex As Exception
            result = If(isMin, " ", "ZZZZZZZZZZ")
        End Try

        Return result
    End Function

    ' Reset and reload data when Refresh button is clicked (FIXED LOGIC)
    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click

        ' 1. প্রথমে কম্বোবক্সের টাইপ করা লেখা ও সিলেকশন ক্লিয়ার করা হচ্ছে (লক খোলার জন্য)
        cmbFromCustomerID.Text = String.Empty
        cmbToCustomerID.Text = String.Empty
        cmbFromAccSetID.Text = String.Empty
        cmbToAccSetID.Text = String.Empty

        cmbFromCustomerID.SelectedIndex = -1
        cmbToCustomerID.SelectedIndex = -1
        cmbFromAccSetID.SelectedIndex = -1
        cmbToAccSetID.SelectedIndex = -1

        ' 2. কন্ট্রোলগুলোকে আবার সচল (Enabled) করা হচ্ছে
        cmbFromCustomerID.Enabled = True
        cmbToCustomerID.Enabled = True
        cmbFromAccSetID.Enabled = True
        cmbToAccSetID.Enabled = True

        ' 3. ডাটাবেজ থেকে ডেটা পুনরায় ফ্রেশ লোড করা হচ্ছে
        LoadCustomerIDsFromView()
        LoadAccountSetsFromView()
        SetDefaultDates()

        ' 4. চেকবক্স রিসেট করা হচ্ছে
        chkDetailJournals.Checked = False

        MessageBox.Show("Data Refreshed and Filters Reset Successfully!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' Pass user selections or dynamic database boundary values to the report viewer window
    Private Sub btnShowReport_Click(sender As Object, e As EventArgs) Handles btnShowReport.Click

        Dim rptForm As New frmReportViewer()

        ' 1. DATE PARAMETERS HANDLING WITH DYNAMIC DEFAULTS
        If dtpFromDate.Checked = False OrElse dtpFromDate.Value = Nothing Then
            rptForm.FromDate = dtpFromDate.Value.Date
        Else
            rptForm.FromDate = dtpFromDate.Value.Date
        End If

        If dtpToDate.Checked = False OrElse dtpToDate.Value = Nothing Then
            rptForm.ToDate = DateTime.Today
        Else
            rptForm.ToDate = dtpToDate.Value.Date
        End If


        ' 2. CUSTOMER ID PARAMETERS HANDLING (Supports Typed Texts)
        Dim fromCustText As String = If(cmbFromCustomerID.SelectedItem IsNot Nothing, cmbFromCustomerID.SelectedItem.ToString().Trim(), cmbFromCustomerID.Text.Trim())
        Dim toCustText As String = If(cmbToCustomerID.SelectedItem IsNot Nothing, cmbToCustomerID.SelectedItem.ToString().Trim(), cmbToCustomerID.Text.Trim())

        If String.IsNullOrEmpty(fromCustText) Then
            rptForm.FromCustomerId = GetFirstOrLastValue("Customer ID", True)
        Else
            rptForm.FromCustomerId = fromCustText
        End If

        If String.IsNullOrEmpty(toCustText) Then
            rptForm.ToCustomerId = GetFirstOrLastValue("Customer ID", False)
        Else
            rptForm.ToCustomerId = toCustText
        End If


        ' 3. ACCOUNT SET ID PARAMETERS HANDLING (Supports Typed Texts)
        Dim fromAccText As String = If(cmbFromAccSetID.SelectedItem IsNot Nothing, cmbFromAccSetID.SelectedItem.ToString().Trim(), cmbFromAccSetID.Text.Trim())
        Dim toAccText As String = If(cmbToAccSetID.SelectedItem IsNot Nothing, cmbToAccSetID.SelectedItem.ToString().Trim(), cmbToAccSetID.Text.Trim())

        If String.IsNullOrEmpty(fromAccText) Then
            rptForm.FromAccSet = GetFirstOrLastValue("Account Set ID", True)
        Else
            rptForm.FromAccSet = fromAccText
        End If

        If String.IsNullOrEmpty(toAccText) Then
            rptForm.ToAccSet = GetFirstOrLastValue("Account Set ID", False)
        Else
            rptForm.ToAccSet = toAccText
        End If


        ' 4. PASSING CHECKBOX VALUE TO CR REPORT VIEWER WINDOW
        rptForm.ShowDetails = chkDetailJournals.Checked


        ' Display the Crystal Report viewer form
        rptForm.ShowDialog()
    End Sub


    ' =========================================================================
    ' MUTUAL EXCLUSION LOGIC (ENABLING/DISABLING CONTROLS DYNAMICALLY)
    ' =========================================================================

    Private Sub cmbFromCustomerID_TextChanged(sender As Object, e As EventArgs) Handles cmbFromCustomerID.SelectedIndexChanged, cmbFromCustomerID.TextChanged
        Dim fromCustText As String = If(cmbFromCustomerID.SelectedItem IsNot Nothing, cmbFromCustomerID.SelectedItem.ToString().Trim(), cmbFromCustomerID.Text.Trim())

        If Not String.IsNullOrEmpty(fromCustText) Then
            cmbFromAccSetID.Enabled = False
            cmbToAccSetID.Enabled = False
        Else
            If String.IsNullOrEmpty(cmbToCustomerID.Text.Trim()) Then
                cmbFromAccSetID.Enabled = True
                cmbToAccSetID.Enabled = True
            End If
        End If
    End Sub

    Private Sub cmbToCustomerID_TextChanged(sender As Object, e As EventArgs) Handles cmbToCustomerID.SelectedIndexChanged, cmbToCustomerID.TextChanged
        Dim toCustText As String = If(cmbToCustomerID.SelectedItem IsNot Nothing, cmbToCustomerID.SelectedItem.ToString().Trim(), cmbToCustomerID.Text.Trim())

        If Not String.IsNullOrEmpty(toCustText) Then
            cmbFromAccSetID.Enabled = False
            cmbToAccSetID.Enabled = False
        Else
            If String.IsNullOrEmpty(cmbFromCustomerID.Text.Trim()) Then
                cmbFromAccSetID.Enabled = True
                cmbToAccSetID.Enabled = True
            End If
        End If
    End Sub

    Private Sub cmbFromAccSetID_TextChanged(sender As Object, e As EventArgs) Handles cmbFromAccSetID.SelectedIndexChanged, cmbFromAccSetID.TextChanged
        Dim fromAccText As String = If(cmbFromAccSetID.SelectedItem IsNot Nothing, cmbFromAccSetID.SelectedItem.ToString().Trim(), cmbFromAccSetID.Text.Trim())

        If Not String.IsNullOrEmpty(fromAccText) Then
            cmbFromCustomerID.Enabled = False
            cmbToCustomerID.Enabled = False
        Else
            If String.IsNullOrEmpty(cmbToAccSetID.Text.Trim()) Then
                cmbFromCustomerID.Enabled = True
                cmbToCustomerID.Enabled = True
            End If
        End If
    End Sub

    Private Sub cmbToAccSetID_TextChanged(sender As Object, e As EventArgs) Handles cmbToAccSetID.SelectedIndexChanged, cmbToAccSetID.TextChanged
        Dim toAccText As String = If(cmbToAccSetID.SelectedItem IsNot Nothing, cmbToAccSetID.SelectedItem.ToString().Trim(), cmbToAccSetID.Text.Trim())

        If Not String.IsNullOrEmpty(toAccText) Then
            cmbFromCustomerID.Enabled = False
            cmbToCustomerID.Enabled = False
        Else
            If String.IsNullOrEmpty(cmbFromAccSetID.Text.Trim()) Then
                cmbFromCustomerID.Enabled = True
                cmbToCustomerID.Enabled = True
            End If
        End If
    End Sub

    Private Sub MainPanel_Paint(sender As Object, e As PaintEventArgs) Handles MainPanel.Paint
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
    End Sub

End Class