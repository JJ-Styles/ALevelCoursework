Imports System.Data.Sql
Imports System.Data.SqlClient
Public Class ProgressStudent
    Structure Data
        Dim Firstname, Surname As String
        Dim Qualification As String
        Dim ResultP, ResultS, ResultW, ResultI As Decimal
        Dim Attempts As Integer
    End Structure
    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        AdminView1.Show()       'Opens the AdminVIew Page
        Me.Close()          'Closes this page
    End Sub
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        StudentLogin.Show()     'Opens the StudentLogin Page
        Me.Close()              'Closes this page
    End Sub
    Private Sub ProgressStudent_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetStudentsNames()          'Calls the procedure GetStudentNames
    End Sub
    Sub GetStudentsNames()
        Dim User() As Data
        Dim i As Integer
        Dim TotalRows As Integer
        Dim Cmd As String = "SELECT Student.FirstName, Student.Surname FROM Student"
        Dim Con As New SqlConnection(ConString)
        Dim Command As New SqlCommand
        Dim reader As SqlDataReader
        i = 0
        Try
            If Con.State = ConnectionState.Closed Then          'Checks is the connection is closed
                Con.Open()              'Sets the connection to open
            End If
            Command = Con.CreateCommand     'Creates the settings for a new SQL Data Reader
            Command.CommandText = "SELECT COUNT(*) FROM Student"        'Sets the command text
            TotalRows = CInt(Command.ExecuteScalar)             'Executes the Command query and stores the integer converted result into the variable
            ReDim User(TotalRows)
            Command = Con.CreateCommand     'Creates the settings for a new SQL Data Reader
            Command.CommandText = Cmd        'Sets the command text
            reader = Command.ExecuteReader    'Executes the command Query and stores the results inside the SQL Data Reader
            Do While reader.Read            'Checks if there is more valuesto be read
                User(i).Firstname = reader.GetValue(0)      'Assigns the value in the reader to the variable
                User(i).Surname = reader.GetValue(1)        'Assigns the value in the reader to the variable
                i += 1          'Increments i by 1
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)         'Sends a error message to the user if an error occurs.
        Finally
            Con.Close()                 'Sets the connection to closed
        End Try
        Try
            For n = 0 To i
                tbxStudentSearch.AutoCompleteCustomSource.Add(User(n).Surname & " " & User(n).Firstname)         'Adds the students names to the source so that when a name is entered into the textbox it predicts who is being searched for
            Next
        Catch ex As Exception
            MsgBox(ex.Message)      'Sends an error message to the user if an error occurs
        End Try
    End Sub
    Private Sub pbxSearch_Click(sender As Object, e As EventArgs) Handles pbxSearch.Click
        Dim Student As String = tbxStudentSearch.Text
        Dim Name(1) As String
        Name = Student.Split(" ")       'Seperates the name into first and surname by the space and places into a array
        GetDataForStudent(Name)         'Calls GetDataForStudent and passes the array name into it.
    End Sub
    Sub GetDataForStudent(ByVal Name() As String)
        Dim User As Data
        Dim Attempt() As Data
        Dim i, n, m As Integer
        Dim TotalRows As Integer
        Dim Con As New SqlConnection(ConString)
        Dim Cmd As String = "SELECT Student.ResultP, Student.ResultS, Student.ResultW, Student.ResultI, Student.Qualification FROM Student WHERE Student.Firstname = '" & Name(1) & "' AND Student.Surname = '" & Name(0) & "'"
        Dim Command As New SqlCommand
        Dim reader As SqlDataReader
        i = 0
        Try
            If Con.State = ConnectionState.Closed Then          'Checks if the connection is closed
                Con.Open()              'Sets the connection to open
            End If
            Command = Con.CreateCommand     'Creates the settings for a new sql command
            Command.CommandText = "SELECT COUNT(*) FROM Test"       'Sets the command text
            TotalRows = CInt(Command.ExecuteScalar)     'Executes the command query and stores the integer converted result in the variable TotalRows
            ReDim Attempt(TotalRows)
            Command = Con.CreateCommand     'Creates the settings for a new sql command
            Command.CommandText = Cmd       'Sets the command text
            reader = Command.ExecuteReader  'Executes the command query and stores the resuls in the sql data reader
            Do While reader.Read    'Checks if the reader has more values to be read
                If IsDBNull(reader.GetValue(0)) Then        'Checks if the value inside the reader is database null
                    User.ResultP = 0            'assigns the value 0 to the variable if the value of the reader is database null
                Else
                    User.ResultP = reader.GetValue(0)       'Assigns the value of the reader to the variable if the reader value isnt a database null
                End If
                If IsDBNull(reader.GetValue(1)) Then        'Checks if the value inside the reader is database null
                    User.ResultS = 0            'assigns the value 0 to the variable if the value of the reader is database null
                Else
                    User.ResultS = reader.GetValue(1)       'Assigns the value of the reader to the variable if the reader value isnt a database null
                End If
                If IsDBNull(reader.GetValue(2)) Then        'Checks if the value inside the reader is database null
                    User.ResultW = 0            'assigns the value 0 to the variable if the value of the reader is database null
                Else
                    User.ResultW = reader.GetValue(2)       'Assigns the value of the reader to the variable if the reader value isnt a database null
                End If
                If IsDBNull(reader.GetValue(3)) Then        'Checks if the value inside the reader is database null
                    User.ResultI = 0            'assigns the value 0 to the variable if the value of the reader is database null
                Else
                    User.ResultI = reader.GetValue(3)       'Assigns the value of the reader to the variable if the reader value isnt a database null
                End If
                If IsDBNull(reader.GetValue(4)) Then        'Checks if the value inside the reader is database null
                    User.Qualification = "N/A"
                Else
                    User.Qualification = reader.GetValue(4)       'Assigns the value of the reader to the variable if the reader value isnt a database null
                End If
            Loop
            If Con.State = ConnectionState.Open Then        'Checks if the connection is open
                Con.Close()         'Sets the connection to closed
                Con.Open()          'Sets the connection to open
            End If
            Command = Con.CreateCommand   'Creates the settings for a new sql command
            Command.CommandText = "SELECT Test.Attempts FROM Test, Student WHERE Student.Surname = '" & Name(0) & "' AND Student.FirstName = '" & Name(1) & "' AND Student.BSCID = Test.BSCID"      'Sets command text
            reader = Command.ExecuteReader          'Executes the command query and stores the results in the sql data reader
            Do While reader.Read        'Checks if the reader has more values to read
                If IsDBNull(reader.GetValue(0)) Then        'Checks if the reader contains a database null value
                    Attempt(i).Attempts = 0             'Assigns the value 0 to the variable if the value inside the reader is database null
                Else
                    Attempt(i).Attempts = reader.GetValue(0)      'Assigns the value inside the reader into the variable
                End If
                i += 1          'Increments i by 1
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)     'Sends a error message to the user if a error occurs.
        Finally
            Con.Close()     'Sets the connection to closed
        End Try
        ManipulateData(User, Attempt)           'Calls the procedure ManipulateData and passes through user and attempt variables
    End Sub
    Sub ManipulateData(ByVal User As Data, ByVal attempt() As Data)
        Try
            tbxQualification.Text = User.Qualification      'Assign the value of the variable to the textbox
            tbxPowerpoint.Text = User.ResultP & "%"      'Assign the value of the variable to the textbox
            tbxSpreadsheet.Text = User.ResultS & "%"      'Assign the value of the variable to the textbox
            tbxWord.Text = User.ResultW & "%"      'Assign the value of the variable to the textbox
            tbxProductivity.Text = User.ResultI & "%"      'Assign the value of the variable to the textbox
            tbxAttemptsP.Text = attempt(0).Attempts      'Assign the value of the variable to the textbox
            tbxAttemptsS.Text = attempt(1).Attempts      'Assign the value of the variable to the textbox
            tbxAttemptsW.Text = attempt(2).Attempts      'Assign the value of the variable to the textbox
            tbxAttemptsI.Text = attempt(3).Attempts      'Assign the value of the variable to the textbox
        Catch ex As Exception
            MsgBox(ex.Message)              'Sends an error message to the user if an error occurs.
        End Try
    End Sub
End Class