Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.IO

Public Class AdminView1

    Structure Data
        Dim Pass As String
    End Structure

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim PassFailRatio As Decimal
        GetPassFailData(PassFailRatio)    'Calls the procedure GetPassFailRatio and passes the variable PassFailRatio into it 
        MsgBox("The Current Pass/Fail Ratio is: " & PassFailRatio & "%") 'Shows the user what the PassFailRatio is through a messagebox
    End Sub

    Sub GetPassFailData(ByRef PassFailRatio As Decimal)
        Dim User() As Data
        Dim TotalRows As Integer
        Dim NoOfStudents As Integer
        Dim i As Integer
        Dim Con As New SqlConnection(ConString)
        Dim Cmd As String = "SELECT Result FROM Test"
        Dim Command As New SqlCommand
        Dim reader As SqlDataReader
        i = 0
        Try
            If Con.State = ConnectionState.Closed Then   'Checks if the connection is closed
                Con.Open()    'Sets the connection to open
            End If
            Command = Con.CreateCommand    'Creates the settings for a new sql command
            Command.CommandText = "SELECT COUNT(*) FROM Student"   'Sets the commend to the string
            NoOfStudents = CInt(Command.ExecuteScalar)     'Executes the command and sets the integer converted value to the variable NoOfStudents
            Command = Con.CreateCommand      'Creates the settings for a new sql command
            Command.CommandText = "SELECT COUNT(*) FROM Test"   'Sets the command string
            TotalRows = CInt(Command.ExecuteScalar)         'Executes the command and sets the integer converted value to the variable TotalRows
            ReDim User(TotalRows - 1) 'redeclares the array user to the size of table in the database
            Command = Con.CreateCommand      'Creates the settings for a new sql command
            Command.CommandText = Cmd       'Sets the command string
            reader = Command.ExecuteReader  'Sest the SqlDataReader to the values that met the command query 
            Do While reader.Read     'Checks to seeif there is more values to be read
                If IsDBNull(reader.GetValue(0)) Then    'Checks whether the value is a database null value
                    User(i).Pass = "failed"   ' Sets the variable if the raeder was containing a database null
                Else
                    User(i).Pass = reader.GetValue(0)   'Otherwise it places the value fromthe reader into the variable
                End If
                i += 1   'Increments the variable by one so that it increases at the same rate each repeat of the loop
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)   'Messagebox to show the user if a error has occured
        Finally
            Con.Close()   'Closed the connection
        End Try
        For n = 0 To (i - 1)
            If User(n).Pass.ToLower = "passed" Then   'checks if the variable is equal to passed. Sets the value inside the variable to all lowe case to catch any case errors.
                PassFailRatio += 1      'if the variable is equal then 1 is added to the current value inside the variable passfailratio
            End If
        Next
        PassFailRatio /= TotalRows  'The variable PassfailRatio is divided by the variable TotalRows to create a fraction
        PassFailRatio *= 100        'Formats the PassFailRatio Variable so that it is a percentage
        PassFailRatio = Math.Round(PassFailRatio, 1) 'Rounds the variable to one decimal place.
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        'Procedure checks the value of the combobox and then opens the correct page dependant on the value inside the combobox.
        If CbxProgress.SelectedItem = "Student" Then
            ProgressStudent.Show()
            Me.Close()
        ElseIf CbxProgress.SelectedItem = "Class: 11E1" Then
            SelectedClass = "11E1"
            ProgressClass.Show()
            Me.Close()
        ElseIf CbxProgress.SelectedItem = "Class: 11E2" Then
            SelectedClass = "11E2"
            ProgressClass.Show()
            Me.Close()
        ElseIf CbxProgress.SelectedItem = "Class: 11E3" Then
            SelectedClass = "11E3"
            ProgressClass.Show()
            Me.Close()
        ElseIf CbxProgress.SelectedItem = "Class: 11W1" Then
            SelectedClass = "11W1"
            ProgressClass.Show()
            Me.Close()
        ElseIf CbxProgress.SelectedItem = "Class: 11W2" Then
            SelectedClass = "11W2"
            ProgressClass.Show()
            Me.Close()
        ElseIf CbxProgress.SelectedItem = "Class: 11W3" Then
            SelectedClass = "11W3"
            ProgressClass.Show()
            Me.Close()
        ElseIf CbxProgress.SelectedItem = "Course by Class" Then
            ProgressCourse.Show()
            Me.Close()
        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        'Procedure checks the value of the combobox and then opens the correct page dependant on the value inside the combobox.
        If CbxReport.SelectedItem = "Class: 11E1" Then
            SelectClassReport = "11E1"
            ReportClass.Show()
            Me.Close()
        ElseIf CbxReport.SelectedItem = "Class: 11E2" Then
            SelectClassReport = "11E2"
            ReportClass.Show()
            Me.Close()
        ElseIf CbxReport.SelectedItem = "Class: 11E3" Then
            SelectClassReport = "11E3"
            ReportClass.Show()
            Me.Close()
        ElseIf CbxReport.SelectedItem = "Class: 11W1" Then
            SelectClassReport = "11W1"
            ReportClass.Show()
            Me.Close()
        ElseIf CbxReport.SelectedItem = "Class: 11W2" Then
            SelectClassReport = "11W2"
            ReportClass.Show()
            Me.Close()
        ElseIf CbxReport.SelectedItem = "Class: 11W3" Then
            SelectClassReport = "11W3"
            ReportClass.Show()
            Me.Close()
        ElseIf CbxReport.SelectedItem = "Course" Then
            ReportCourse.Show()
            Me.Close()
        End If
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Me.Show() 'Calls itslef and refreshes the page
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        StudentLogin.Show()   'Opens the studentLogin Page
        Me.Close()            'Closes this page
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        AddData.Show() 'Opens the AddData Page
        Me.Close()      'Closes this page
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PopUP.Show()        'Opens the PopUp page
    End Sub
End Class