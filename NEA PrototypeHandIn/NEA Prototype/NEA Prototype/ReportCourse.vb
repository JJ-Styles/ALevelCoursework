Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class ReportCourse

    Structure Data
        Dim FirstName, Surname, Passed As String
        Dim Qualification As String
    End Structure

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        AdminView1.Show()       'Opens the AdminView Page
        Me.Close()          'Closes this page
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        StudentLogin.Show()         'Opens the StudentLogin Page
        Me.Close()          'Closes this page
    End Sub

    Private Sub ReportCourse_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetData()           'Calls the GetData Procedure
    End Sub

    Sub GetData()
        Dim User() As Data
        Dim i, n, m As Integer
        Dim TotalRows As Integer
        Dim Cmd As String = "SELECT Student.FirstName, Student.Surname, Student.Qualification, Test.Result FROM Student, Test WHERE Student.BSCID = Test.BSCID ORDER BY Surname"
        Dim Con As New SqlConnection(ConString)
        Dim Command As New SqlCommand
        Dim reader As SqlDataReader
        Dim PassFailRatio As Decimal
        i = 0       'Assigns the value 0 to i 
        Try
            If Con.State = ConnectionState.Closed Then      'Checks if the connection state is closed
                Con.Open()          'Opens the sql database connection
            End If
            Command = Con.CreateCommand             'Creates the settings for a new sql command
            Command.CommandText = "SELECT COUNT(*) FROM Test"       'Sets the command text
            TotalRows = CInt(Command.ExecuteScalar)         'Executes the command query and stores the integer converted result in the variable TotalRows
            ReDim User(TotalRows)
            Command = Con.CreateCommand         'Creates the settings for a new sql command 
            Command.CommandText = Cmd           'Sets the command text
            reader = Command.ExecuteReader      'Executes the command query and stores the results in the sql data reader 
            Do While reader.Read            'Checks if there is more values to be read
                If User(n).Surname <> reader.GetValue(1) Then           'Checks if the the variable is not equal to value inside the reader
                    User(i).FirstName = reader.GetValue(0)              'Assigns the value inside the reader to the variable
                    User(i).Surname = reader.GetValue(1)              'Assigns the value inside the reader to the variable
                    If IsDBNull(reader.GetValue(2)) Then                'Checks if the reader contains a database null value
                        User(i).Qualification = "N/A"                   'Assigns the value N/A if the reader contained a database null value
                    Else
                        User(i).Qualification = reader.GetValue(2)              'Assigns the value inside the reader to the variable
                    End If
                    n = i                           'Assign the value of the variable i to n
                    i += 1                      'Increments i by 1
                End If
                If IsDBNull(reader.GetValue(3)) Then                'Checks if the reader contains a database null value
                    User(m).Passed = "Failed"               'Assigns failed to the variable if the reader contained a database null value
                Else
                    User(m).Passed = reader.GetValue(3)              'Assigns the value inside the reader to the variable
                End If
                m += 1                  'Increments m by 1
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)                 'Sends an error message to the user if a error occurs
        Finally
            Con.Close()                 'Sets the connection state to closed
        End Try
        Try
            For n = 0 To (m - 1)
                If User(n).Passed.ToLower = "passed" Then           'Checks if the variable is equal to passed
                    PassFailRatio += 1                  'Increments the variable by 1 if the IF statement is correct
                End If
            Next
            PassFailRatio /= TotalRows                  'Divides the variable by the variable TotalRows to create a fraction
            PassFailRatio *= 100                        'Multiplies the variable by 100 so that it is in percentage format
            PassFailRatio = Math.Round(PassFailRatio, 1)        'Rounds the variable to 1 decimal place
            tbxPassFailRatio.Text = PassFailRatio & "%"         'Assigns the value inside the variable to the textbox with a following percentage sign
            For n = 0 To (i - 1)                    'Keeps the data in alignment so that each student is placed with the corrected qualification
                LbxName.Items.Add(User(n).Surname & " " & User(n).FirstName)                'Adds the names of each student to the listbox name 
                LbxQualification.Items.Add(User(n).Qualification)                           'Adds the qualification of each student to the listbox Qualification
            Next
        Catch ex As Exception
            MsgBox(ex.Message)                  'Sends an error message to the user if a error occurs.
        End Try
    End Sub

End Class