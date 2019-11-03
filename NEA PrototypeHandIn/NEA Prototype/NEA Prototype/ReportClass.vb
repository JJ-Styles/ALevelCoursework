Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class ReportClass

    Structure Data
        Dim FirstName, Surname, Passed As String
        Dim Qualification, BSCID As String
    End Structure

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        AdminView1.Show()   'Opens the AdminView Page
        Me.Close()      'Closes this page
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        StudentLogin.Show()     'Opens the StudentLogin Page
        Me.Close()          'Closes this Page
    End Sub

    Private Sub ReportClass_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Cmd As String
        Cmd = "SELECT Student.FirstName, Student.Surname, Student.Qualification, Test.Result, Student.BSCID FROM Student, Test WHERE Student.ClassName = '" & SelectClassReport & "' AND Student.BSCID = Test.BSCID"        'Assigns the string to the Cmd variable
        GetData(Cmd)        'Calls the GetData Procedure and Passes the cmd variable into it.
    End Sub

    Sub GetData(ByVal Cmd As String)
        Dim User() As Data
        Dim i, n, m As Integer
        Dim TotalRows As Integer
        Dim Con As New SqlConnection(ConString)
        Dim Command As New SqlCommand
        Dim reader As SqlDataReader
        Dim PassFailRatio As Decimal
        i = 0
        Try
            If Con.State = ConnectionState.Closed Then      'Checks if the Connection is closed
                Con.Open()              'Sets the connection to open
            End If
            Command = Con.CreateCommand         'Creates the setting for a new sql command 
            Command.CommandText = "SELECT COUNT(*) FROM Test WHERE ClassName = '" & SelectClassReport & "'"           'Sets the command text
            TotalRows = CInt(Command.ExecuteScalar)     'Executes the command query and stores the integer converted result to the variable
            ReDim User(TotalRows)
            Command = Con.CreateCommand         'Creates the setting for a new sql command
            Command.CommandText = Cmd           'Sets the command text
            reader = Command.ExecuteReader      'Executes the command query amd stoes the results in the sql data reader
            Do While reader.Read            'Checks if the reader hads more values to be read
                If User(n).BSCID <> reader.GetValue(4) Then     'Checks if the BSCID is the same as the previous
                    User(i).BSCID = reader.GetValue(4)          'Assigns the value in the reader to the variable
                    User(i).FirstName = reader.GetValue(0)          'Assigns the value in the reader to the variable
                    User(i).Surname = reader.GetValue(1)          'Assigns the value in the reader to the variable
                    If IsDBNull(reader.GetValue(2)) Then           'Checks if the reader contains a database null value
                        User(i).Qualification = "N/A"               'Assigns the value N/A if the value inside the reader is a database null
                    Else
                        User(i).Qualification = reader.GetValue(2)          'Assigns the value in the reader to the variable
                    End If
                    n = i               'Assign the value i to n
                    i += 1              'Increments i by 1
                End If
                If IsDBNull(reader.GetValue(3)) Then           'Checks if the reader contains a database null value
                    User(m).Passed = "Failed"               'Assigns the value N/A if the value inside the reader is a database null
                Else
                    User(m).Passed = reader.GetValue(3)          'Assigns the value in the reader to the variable
                End If
                m += 1              'Increments m by 1
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)         'Sends an error to the user if a error occurs
        Finally
            Con.Close()         'Sets the connection to close
        End Try
        Try
            For l = 0 To (m - 1)
                If User(l).Passed.ToLower = "passed" Then       'Checks if the variable in lower case is equal to passed
                    PassFailRatio += 1      'Increments the variable by 1 if the IF statement is true
                End If
            Next
            PassFailRatio /= TotalRows      'Divides the variable by TotalRows 
            PassFailRatio *= 100            'Multiplies the variable by 100
            PassFailRatio = Math.Round(PassFailRatio, 1)        'Rounds the variable to 1 decimal place
            tbxPassFailRatio.Text = PassFailRatio & "%"         'Puts the value of the variable in the textbox
            For l = 0 To n
                LbxName.Items.Add(User(l).Surname & " " & User(l).FirstName)        'Adds the fullname to the listbox 
                LbxQualification.Items.Add(User(l).Qualification)           'Adds the qualifiaction of each student in the class to the listbox
            Next
        Catch ex As Exception
            MsgBox(ex.Message)          'Sends an error message to the user if an error occurs.
        End Try
    End Sub

End Class