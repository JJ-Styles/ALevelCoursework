Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Configuration.ConfigurationSettings
Public Class AdminLogin
    Dim User() As Admin
    Dim i As Integer
    Dim TotalRows As Integer

    Structure Admin
        Dim Username As String
        Dim Password As String
    End Structure

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        StudentLogin.Show()   'Opens the StudentLogin Page
        Me.Close()            'Closes this page
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Incorrect As Boolean = True
        Dim password As String
        Dim encryption As String

        password = txtPassword.Text 'Takes the data  inside the password textbox and stores in the password variable.
        encryption = Encrypted(password) 'Calls the Encrypted procedure and passes password into, storing the returned in the encryption variable.
        GetData()   'Calls the GetData Procedure 
        Try
            For n = 0 To (TotalRows - 1)
                If User(n).Username = txtUsername.Text Then      'Checks if the value from the database is the same as the value in the textbox
                    If User(n).Password = encryption Then        'Checks if the value from the database is the same as the value in the textbox
                        Incorrect = False                        'Sets the variable incorrect to true so that the message box isnt displayed
                        AdminView1.Show()                        'Opens the Admin View Page
                        Me.Close()                               'Closes this page
                    End If
                End If
            Next
            If Incorrect Then
                MsgBox("Username, Password are Incorrect, Please Try Again.")  'Displays a error message to the user to indicate that the values are incorrected
            End If
        Catch ex As Exception
            MsgBox(ex.Message)   'Send the error message to the user
        End Try
        txtPassword.Clear()  'Clears the password textbox
    End Sub

    Sub GetData()
        Dim Con As New SqlConnection(ConString)
        Dim Cmd As String = "SELECT Username, Password FROM Admin"
        Dim Command As New SqlCommand
        Dim reader As SqlDataReader
        i = 0
        Try
            If Con.State = ConnectionState.Closed Then  'Checks to see if the connection is closed
                Con.Open()   'Sets the connection to open
            End If
            Command = Con.CreateCommand   'Creates the settings for a new sql command
            Command.CommandText = "SELECT COUNT(*) FROM Admin"  'Sets the command string for the query
            TotalRows = CInt(Command.ExecuteScalar)   'Executes the command query and stores the integer converted result to the variable TotalRows
            ReDim User(TotalRows)
            Command = Con.CreateCommand   'Creates the settings for a new sql command
            Command.CommandText = Cmd  'Sets the command string for the query
            reader = Command.ExecuteReader  'Executes the command and places the result within the SQL data Reader
            Do While reader.Read   'Checks to see if there is another piece of data to be read
                User(i).Username = reader.GetValue(0)   'Takes the data from the reader and stores it in the varaible
                User(i).Password = reader.GetValue(1)   'Takes the data from the reader and stores it in the varaible
                i += 1 'Increments the variable for each iteration of the loop
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)    'Displays a message box to the user if there is an error
        Finally
            Con.Close()    'Sets the connection to close
        End Try
    End Sub

    Function Encrypted(ByRef password As String)
        Dim word As String = password
        Dim Encryption As String
        Dim recursion As Integer

        word = AddWord(word)
        Encryption = railway(word, recursion)
        Return Encryption
    End Function 'See AddData, Encrypted Procedure

    Function AddWord(ByRef word As String)
        Dim Letter As Char
        Do Until word.Length >= 30
            For i = 1 To word.Length
                Letter = Mid(word.ToLower, i, 1)
                Select Case Letter
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
    End Function  'See AddData, AddWord Procedure

    Function railway(ByRef word As String, ByRef Recursion As Integer)
        Dim layer1(14) As Char
        Dim layer2(14) As Char
        Dim n As Integer = 0
        Dim Encryption As String

        For i = 1 To 29 Step 2
            layer1(n) = Mid(word, i, 1)
            n += 1
        Next
        n = 0
        For i = 2 To 30 Step 2
            layer2(n) = Mid(word, i, 1)
            n += 1
        Next
        For i = 0 To 14
            Encryption += layer1(i)
        Next
        For i = 0 To 14
            Encryption += layer2(i)
        Next
        Do Until Recursion >= 10
            Recursion += 1
            Encryption = railway(Encryption, Recursion)
        Loop
        Return Encryption
    End Function  'See AddData, railway Procedure

    Private Sub txtPassword_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPassword.KeyDown
        If e.KeyCode = Keys.KeyCode.Enter Then  'Checks if the enter buttonhas been pressed and activates the login button
            Button1_Click(sender, e)
        End If
    End Sub
End Class