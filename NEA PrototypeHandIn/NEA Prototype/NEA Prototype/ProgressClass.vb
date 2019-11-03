Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class ProgressClass
    Structure Data
        Dim Firstname As String
        Dim Surname As String
        Dim ResultP As Decimal
        Dim ResultS As Decimal
        Dim ResultW As Decimal
        Dim ResultI As Decimal
    End Structure

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        AdminView1.Show() 'Opens the AdminView Page
        Me.Close()        'Close this page
    End Sub

    Private Sub ProgressClass_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tbxClassname.Text = SelectedClass  'Sets the textboxes text to the selected text
        GetData()       'Calls the GetData procedure
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        StudentLogin.Show()    'Opens the StudentLogin page
        Me.Close()              'Closes this page
    End Sub

    Sub GetData()
        Dim User() As Data
        Dim i As Integer
        Dim TotalRows As Integer
        Dim Con As New SqlConnection(ConString)
        Dim Cmd As String = "SELECT Student.FirstName, Student.Surname, Student.ResultP, Student.ResultS, Student.ResultW, Student.ResultI, Class.Invigilator, Class.Trainer FROM Student, Class WHERE Student.ClassName = '" & SelectedClass & "' AND Student.ClassName = Class.ClassName ORDER BY Student.Surname"
        Dim Command As New SqlCommand
        Dim reader As SqlDataReader
        Dim Invigilator As String = Nothing
        Dim Trainer As String = Nothing
        i = 0
        Try
            If Con.State = ConnectionState.Closed Then          'Checks if the connection is closed
                Con.Open()                              'Sets the connection to open
            End If
            Command = Con.CreateCommand             'Creates the settings for a new SQL Command
            Command.CommandText = "SELECT COUNT(*) FROM Student"           'Sets the command text
            TotalRows = CInt(Command.ExecuteScalar)         'Executes the command query and stores the integer converter result into the TotalRows variable
            ReDim User(TotalRows)
            Command = Con.CreateCommand             'Creates the settings for a new SQL Command
            Command.CommandText = Cmd               'Sets the command text
            reader = Command.ExecuteReader          'Executes the command query and stores the results inside the sql data reader
            Do While reader.Read                    'Checks if there is still data to read
                User(i).Firstname = reader.GetValue(0)      'Gets the value from the reader and stores the value inside the variable
                User(i).Surname = reader.GetValue(1)      'Gets the value from the reader and stores the value inside the variable
                If IsDBNull(reader.GetValue(2)) Then      'Checks if the value inside the reader is a database null
                    User(i).ResultP = 0                   'If the value is a database null it stores  into the variable
                Else
                    User(i).ResultP = reader.GetValue(2)      'Gets the value from the reader and stores the value inside the variable
                End If
                If IsDBNull(reader.GetValue(3)) Then      'Checks if the value inside the reader is a database null
                    User(i).ResultS = 0                   'If the value is a database null it stores  into the variable
                Else
                    User(i).ResultS = reader.GetValue(3)      'Gets the value from the reader and stores the value inside the variable
                End If
                If IsDBNull(reader.GetValue(4)) Then      'Checks if the value inside the reader is a database null
                    User(i).ResultW = 0                   'If the value is a database null it stores  into the variable
                Else
                    User(i).ResultW = reader.GetValue(4)      'Gets the value from the reader and stores the value inside the variable
                End If
                If IsDBNull(reader.GetValue(5)) Then      'Checks if the value inside the reader is a database null
                    User(i).ResultI = 0                   'If the value is a database null it stores  into the variable
                Else
                    User(i).ResultI = reader.GetValue(5)      'Gets the value from the reader and stores the value inside the variable
                End If
                Invigilator = reader.GetValue(6)      'Gets the value from the reader and stores the value inside the variable
                Trainer = reader.GetValue(7)      'Gets the value from the reader and stores the value inside the variable
                i += 1
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)                 'Sends a error message to the user if an error occurs
        Finally
            Con.Close()             'Sets the connection to close
        End Try
        Try
            For n = 0 To i
                FormatPercent(User(n).ResultP)          'Formats the variable into a percentage
                FormatPercent(User(n).ResultS)          'Formats the variable into a percentage
                FormatPercent(User(n).ResultW)          'Formats the variable into a percentage
                FormatPercent(User(n).ResultI)          'Formats the variable into a percentage
            Next
            tbxInvigilator.Text = Invigilator           'Sets the invigilator textbox to the data stored in the inivigilator variable 
            tbxTrainer.Text = Trainer                   'Sets the trainer textbox to the data stored in the trainer variable
            For n = 0 To (i - 1)
                LbxName.Items.Add(User(n).Surname & " " & User(n).Firstname)        'Adds the fullname of the students to the listbox name
                LbxPowerpoint.Items.Add(User(n).ResultP & "%")                      'Adds the Powerpoint result of each student to the listbox Powerpoint
                LbxSpreadsheet.Items.Add(User(n).ResultS & "%")                     'Adds the spreadsheet result of each student to the listbox spreadsheet
                LbxWord.Items.Add(User(n).ResultW & "%")                            'Adds the word result of each student to the listbox word
                LbxProductivity.Items.Add(User(n).ResultI & "%")                    'Adds the Productivity resultof each student to the listbox Productivity
            Next
        Catch ex As Exception
            MsgBox(ex.Message)      'Sends and error message to the user if one occurs
        End Try
    End Sub

End Class