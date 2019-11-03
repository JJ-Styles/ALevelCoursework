Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Configuration.ConfigurationSettings
Public Class AddData
    Private Sub btnSave_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Reset()             'Calls Reset procedure
    End Sub
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles pbxSearch.Click
        Reset()                         ' Calls Reset Procedure
        If cbxTable.SelectedItem = "Admin" Then         ' Checks the combobox to see if admin has been selected
            Admin()                                     ' Calls Admin Procedure
        ElseIf cbxTable.SelectedItem = "Student" Then    ' Checks the combobox to see if Student has been selected
            Student()                                     ' Calls Student Procedure
        ElseIf cbxTable.SelectedItem = "Class" Then      ' Checks the combobox to see if class has been selected
            Classes()                                     ' Calls Classes Procedure
        ElseIf cbxTable.SelectedItem = "Unit" Then       ' Checks the combobox to see if unit has been selected
            Unit()                                     ' Calls Unit Procedure
        ElseIf cbxTable.SelectedItem = "Test" Then       ' Checks the combobox to see if test has been selected
            Test()                                     ' Calls Test Procedure
        End If
    End Sub
    Sub Admin()
        lblColumn1.Show()                           'Shows the top label
        lblColumn1.Text = "Username:"               'Sets the top label name
        lblColumn2.Show()                           'Shows the next label
        lblColumn2.Text = "Password:"               'Sets the next label name
        lblColumn3.Show()                           'Shows the next label
        lblColumn3.Text = "Name:"                   'Sets the next label name
        tbxColumn1.MaxLength = 30                   'Sets a max length to the top textbox
        tbxColumn1.Show()                           'Shows the top textbox
        tbxColumn2.MaxLength = 30                   'Sets a max length to the next textbox
        tbxColumn2.Show()                           'Shows the next textbox
        tbxColumn2.UseSystemPasswordChar = True     'Sets the textbox to show the password character
        tbxColumn3.MaxLength = 30                   'Sets a max length to the next textbox
        tbxColumn3.Show()                           'Shows the next textbox
    End Sub
    Sub Student()
        lblColumn1.Show()                           'Shows the top label
        lblColumn1.Text = "BSCID:"                  'Sets the top label name
        lblColumn2.Show()                           'Shows the next label
        lblColumn2.Text = "FirstName:"              'Sets the next label name
        lblColumn3.Show()                           'Shows the next label
        lblColumn3.Text = "Surname:"                'Sets the next label name
        lblColumn4.Show()                           'Shows the next label
        lblColumn4.Text = "Username"                'Sets the next label name
        lblColumn5.Show()                           'Shows the next label
        lblColumn5.Text = "Password"                'Sets the next label name
        lblColumn6.Show()                           'Shows the next label
        lblColumn6.Text = "Registration Date:"      'Sets the next label name
        lblColumn7.Show()                           'Shows the next label
        lblColumn7.Text = "Completion Date"         'Sets the next label name
        lblColumn8.Show()                           'Shows the next label
        lblColumn8.Text = "Certificate Date:"       'Sets the next label name
        lblColumn9.Show()                           'Shows the next label
        lblColumn9.Text = "Class Name:"             'Sets the next label name
        tbxColumn1.Show()                           'Shows the top textbox
        tbxColumn1.MaxLength = 12                   'Sets a max length to the top textbox
        tbxColumn2.Show()                           'Shows the next textbox
        tbxColumn2.MaxLength = 30                   'Sets a max length to the next textbox
        tbxColumn3.Show()                           'Shows the next textbox
        tbxColumn3.MaxLength = 30                   'Sets a max length to the next textbox
        tbxColumn4.Show()                           'Shows the next textbox
        tbxColumn4.MaxLength = 30                   'Sets a max length to the next textbox
        tbxColumn5.Show()                           'Shows the next textbox
        tbxColumn5.MaxLength = 30                   'Sets a max length to the next textbox
        tbxColumn5.UseSystemPasswordChar = True     'Sets the textbox to show the password character
        dtpColumn6.Show()                           'Shows the next date picker
        dtpColumn7.Show()                           'Shows the next date picker
        dtpColumn8.Show()                           'Shows the next date picker
        tbxColumn9.Show()                           'Shows the next textbox
        tbxColumn9.MaxLength = 4                    'Sets a max length to the next textbox
    End Sub
    Sub Classes()
        lblColumn1.Show()                           'Shows the top label
        lblColumn1.Text = "Class Name:"             'Sets the top label name
        lblColumn2.Show()                           'Shows the next label
        lblColumn2.Text = "Invigilator:"            'Sets the next label name
        lblColumn3.Show()                           'Shows the next label
        lblColumn3.Text = "Trainer:"                'Sets the next label name
        tbxColumn1.Show()                           'Shows the top textbox
        tbxColumn1.MaxLength = 4                    'Sets a max length to the top TextBox
        tbxColumn2.Show()                           'Shows the next textbox
        tbxColumn2.MaxLength = 30                   'Sets a max length to the next textbox
        tbxColumn3.Show()                           'Shows the next textbox
        tbxColumn3.MaxLength = 30                   'Sets a max length to the next textbox
    End Sub
    Sub Unit()
        lblColumn1.Show()                           'Shows the top label
        lblColumn1.Text = "Unit Code:"              'Sets the top label name
        lblColumn2.Show()                           'Shows the next label
        lblColumn2.Text = "Description:"            'Sets the next label name
        lblColumn3.Show()                           'Shows the next label
        lblColumn3.Text = "Syllabus:"               'Sets the next label name
        tbxColumn1.Show()                           'Shows the top textbox
        tbxColumn1.MaxLength = 4                    'Sets a max length to the top textbox
        tbxColumn2.Show()                           'Shows the next textbox
        tbxColumn2.MaxLength = 20                   'Sets a max length to the next textbox
        tbxColumn3.Show()                           'Shows the next textbox
        tbxColumn3.MaxLength = 4                    'Sets a max length to the next textbox
    End Sub
    Sub Test()
        lblColumn1.Show()                           'Shows the top label
        lblColumn1.Text = "Test Paper:"             'Sets the top label name
        lblColumn2.Show()                           'Shows the next label
        lblColumn2.Text = "Result:"                 'Sets the next label name
        lblColumn3.Show()                           'Shows the next label
        lblColumn3.Text = "Attempts:"               'Sets the next label name
        lblColumn4.Show()                           'Shows the next label
        lblColumn4.Text = "Test Date"               'Sets the next label name
        lblColumn5.Show()                           'Shows the next label
        lblColumn5.Text = "Start Date:"             'Sets the next label name
        lblColumn6.Show()                           'Shows the next label
        lblColumn6.Text = "BCSID"                   'Sets the next label name
        lblColumn7.Show()                           'Shows the next label
        lblColumn7.Text = "Unit Code"               'Sets the next label name
        tbxColumn1.Show()                           'Shows the top textbox
        tbxColumn1.MaxLength = 6                    'Sets a max length to the top textbox
        tbxColumn2.Show()                           'Shows the next textbox
        ToolTip1.Active = True                      'Sets the tool tip to actuve so that it appears to the user
        tbxColumn2.MaxLength = 6                    'Sets a max length to the next textbox
        tbxColumn3.Show()                           'Shows the next textbox
        tbxColumn3.MaxLength = 30                   'Sets a max length to the next textbox
        dtpColumn4.Show()                           'Shows the next date picker
        dtpColumn5.Show()                           'Shows the next date picker
        tbxColumn6.Show()                           'Shows the next date picker
        tbxColumn7.Show()                           'Shows the next date picker
    End Sub
    Sub Reset()                     'Sets all screen textboxes, datepickers and labels back to default settings
        lblColumn1.Hide()
        lblColumn2.Hide()
        lblColumn3.Hide()
        lblColumn4.Hide()
        lblColumn5.Hide()
        lblColumn6.Hide()
        lblColumn7.Hide()
        lblColumn8.Hide()
        lblColumn9.Hide()
        lblColumn10.Hide()
        tbxColumn1.Hide()
        tbxColumn1.Clear()
        tbxColumn2.Hide()
        tbxColumn2.UseSystemPasswordChar = False
        tbxColumn2.Clear()
        tbxColumn3.Hide()
        tbxColumn3.Clear()
        tbxColumn4.Hide()
        tbxColumn4.Clear()
        tbxColumn5.Hide()
        tbxColumn5.Clear()
        tbxColumn5.UseSystemPasswordChar = False
        tbxColumn6.Hide()
        tbxColumn6.Clear()
        tbxColumn7.Hide()
        tbxColumn7.Clear()
        tbxColumn8.Hide()
        tbxColumn8.Clear()
        tbxColumn9.Hide()
        tbxColumn9.Clear()
        tbxColumn10.Hide()
        tbxColumn10.Clear()
        dtpColumn4.Hide()
        dtpColumn5.Hide()
        dtpColumn6.Hide()
        dtpColumn7.Hide()
        dtpColumn8.Hide()
        ToolTip1.Active = False
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim Cmd As String
        Dim Password As String
        If cbxTable.SelectedItem = "Admin" Then 'Checks if the comobox was set to admin
            Password = tbxColumn2.Text    ' Picks up what is inside the texbox and places it within the variable password
            Password = Encrypted(Password) 'sets pasword equal to the return value of the procedure where the value that pasword was previously gets set passed into it
            Cmd = "INSERT INTO Admin (Username, Password, Name) VALUES ('" & tbxColumn1.Text & "','" & Password & "','" & tbxColumn3.Text & "')" 'Sets the Cmd to this string, Picking up what is in each textbox and the password variable
            AddToTable(Cmd) 'Calls the AddToTable procedure and sends the Cmd to it as well.
        ElseIf cbxTable.SelectedItem = "Student" Then  'Checks if the combobox was set to student
            Password = tbxColumn5.Text 'Assigns the value inside the texboxe to the password variable
            Password = Encrypted(Password) 'Calss the Encrypted procedure and passes the password variable into it. once the function is complete a value is returned and placed inside the password variable.
            Cmd = "INSERT INTO Student (BSCID, FirstName, Surname, Username, Password, RegDate, CompDate, CertificateDate, ClassName) VALUES ('" & tbxColumn1.Text & "','" & tbxColumn2.Text & "','" & tbxColumn3.Text & "','" & tbxColumn4.Text & "','" & Password & "','" & (dtpColumn6.Text) & "','" & (dtpColumn7.Text) & "','" & (dtpColumn4.Text) & "','" & tbxColumn9.Text & "')" 'Sets the string value to the cmd, Placing in the values inside the textboxes, datepickers and the password variable
            AddToTable(Cmd)   'Calls the AddToTable procedure and passes the Cmd into it.
        ElseIf cbxTable.SelectedItem = "Class" Then 'Checks if the comobox was set to Class
            Cmd = "INSERT INTO Class (ClassName, Invigilator, Trainer) VALUES ('" & tbxColumn1.Text & "','" & tbxColumn2.Text & "','" & tbxColumn3.Text & "')" 'Sets the Cmd to the string and collects the data inside the textboxes
            AddToTable(Cmd)   'Calls the AddToTable procedure and passes the Cmd into it.
        ElseIf cbxTable.SelectedItem = "Unit" Then 'Checks if the comobox was set to Unit
            Cmd = "INSERT INTO Unit (UnitCode, Description, Syllabus) VALUES ('" & tbxColumn1.Text & "','" & tbxColumn2.Text & "','" & tbxColumn3.Text & "')"  'Sets the Cmd to the string and collects the data inside the textboxes
            AddToTable(Cmd)   'Calls the AddToTable procedure and passes the Cmd into it.
        ElseIf cbxTable.SelectedItem = "Test" Then 'Checks if the comobox was set to Test
            Cmd = "INSERT INTO Test (TestPaper, Result, Attempts, TestDate, StartDate,BSCID, UnitCode) VALUES ('" & tbxColumn1.Text & "','" & tbxColumn2.Text & "','" & tbxColumn3.Text & "','" & dtpColumn4.Text & "','" & dtpColumn5.Text & "','" & tbxColumn6.Text & "','" & tbxColumn7.Text & "')"    'Sets the Cmd to the string and collects the data inside the textboxes and datepickers
            AddToTable(Cmd)   'Calls the AddToTable procedure and passes the Cmd into it.
        End If
        Reset()   'Calls the reset Procedure
    End Sub
    Sub AddToTable(ByVal cmd As String)
        'The Declarations set up the states required to connect to the database
        Dim Con As New SqlConnection(ConString)
        Dim Command As New SqlCommand(cmd, Con)
        Dim Writer As New SqlDataAdapter

        Try
            If Con.State = ConnectionState.Closed Then   'Checks if the the connection is closed
                Con.Open()     'Sets the connection to open
            ElseIf Con.State <> ConnectionState.Closed Then   'Checks to see if the connection is open 
                Con.Close()      'Sets the connection to close 
                Con.Open()       'Sets the connection to open 
            End If
            Command = Con.CreateCommand     'Creates the command settings
            Command.CommandText = cmd       'Sets the command to Cmd string
            Command.ExecuteNonQuery()       'Executes the the command
            MsgBox("Data Added to Table.")  'Message box tells the user that it has added the data 
        Catch ex As Exception
            MsgBox(ex.Message)              'Error handling message box shows the user the error if one occurs
        Finally
            Con.Close()                     'Sets the connection to close
        End Try

    End Sub
    Function Encrypted(ByRef password As String)
        Dim word As String = password
        Dim Encryption As String
        Dim recursion As Integer

        word = AddWord(word)    'Calls the AddWord procedure and pass the value word into it. Once the procedure is completed a value is return and stored inside the word variable.
        Encryption = railway(word, recursion) 'Calls the railway procedure and passes word and recursion into it. Once the procedure is completed the returned value gets stored in the encryption variable.
        Return Encryption 'returns the value stord in the variable encryption.
    End Function
    Function AddWord(ByRef word As String)
        Dim Letter As Char
        Do Until word.Length >= 30     'checks to see if the letter is equal to or less than 30 and repeats until it is complete.
            For i = 1 To word.Length   'Cycles through the for statement making i = 1 to the length of the word
                Letter = Mid(word.ToLower, i, 1)  'Sets the letter variable to the character selected by the mid function.
                Select Case Letter  'Checks what value is inside the letter variable and adds another letter to the word variable dependant on what letter has been selected.
                    Case "a"
                        word += "q"
                    Case "b"
                        word += "6"
                    Case "c"
                        word += "h"
                    Case "d"
                        word += "8"
                    Case "e"
                        word += "9"
                    Case "f"
                        word += "h"
                    Case "g"
                        word += "4"
                    Case "h"
                        word += "£"
                    Case "i"
                        word += "j"
                    Case "j"
                        word += "5"
                    Case "k"
                        word += "z"
                    Case "l"
                        word += "7"
                    Case "m"
                        word += "p"
                    Case "n"
                        word += "l"
                    Case "o"
                        word += "2"
                    Case "p"
                        word += "0"
                    Case "q"
                        word += "z"
                    Case "r"
                        word += "c"
                    Case "s"
                        word += "1"
                    Case "t"
                        word += "m"
                    Case "u"
                        word += "s"
                    Case "v"
                        word += "2"
                    Case "w"
                        word += "3"
                    Case "x"
                        word += "a"
                    Case "y"
                        word += "0"
                    Case "z"
                        word += "6"
                End Select
            Next
        Loop
        Return word
    End Function
    Function railway(ByRef word As String, ByRef Recursion As Integer)
        Dim layer1(14) As Char
        Dim layer2(14) As Char
        Dim n As Integer = 0
        Dim Encryption As String

        For i = 1 To 29 Step 2 'For statement that picks up every second letter
            layer1(n) = Mid(word, i, 1) 'Mid function picks up all the letters that are in an odd position within the word and stores it in the array
            n += 1   'increments the value of n so that the array is increased with each cycle of the for loop
        Next
        n = 0     'Reset n so that it can be used again
        For i = 2 To 30 Step 2   'For statement that picks up every second letter
            layer2(n) = Mid(word, i, 1)  'Mid function picks up all the letters that are in an even position within the word and stores it in the array
            n += 1   'increments the value of n so that the array is increased with each cycle of the for loop
        Next
        For i = 0 To 14
            Encryption += layer1(i)  'adds the character inside the array to the string encryption
        Next
        For i = 0 To 14
            Encryption += layer2(i)  'adds the character inside the array to the string encryption
        Next
        Do Until Recursion >= 10 'Checks to see the value of the variable recursion is greater than or equal to 10
            Recursion += 1 'If the recursion isnt greater than or less than it adds one to the recursion
            Encryption = railway(Encryption, Recursion) 'Calls the procedure again to scramble the letters even more.
        Loop
        Return Encryption   'Returns the string inside the variable Encryption
    End Function
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        StudentLogin.Show()    'Opens the StudentLogin Page
        Me.Close()             'Closes this page
    End Sub
    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        AdminView1.Show()       'Opens the AdminView Page
        Me.Close()              'Closes this page
    End Sub
End Class