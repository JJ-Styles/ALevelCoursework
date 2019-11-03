Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Configuration.ConfigurationSettings
Public Class StudentLogin
    Dim User() As Student
    Dim i As Integer
    Dim TotalRows As Integer
    Structure Student
        Dim Username As String
        Dim Password As String
    End Structure
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Incorrect As Boolean = True
        Dim password As String
        Dim encryption As String
        StudentUser = txtUsername.Text          'Assigns the value inside the textbox to the variable
        password = txtPassword.Text.ToLower     'Assigns the value inside the textbox to the variable
        encryption = Encrypted(password)        'Calls the procedure encyrption and passes through the variable password, storing the output in the variable encryption
        GetData()                   'Calls the GetData Procedure
        Try
            For n = 0 To (TotalRows - 1)
                If User(n).Username = txtUsername.Text Then             'Checks if the value inside the variable is the same as the textbox
                    If User(n).Password = encryption Then             'Checks if the value inside the variable is the same as the textbox
                        Incorrect = False               'Assigns the value false to the variable
                        StudentView.Show()              'Opens the page StudentView
                        Me.Close()                      'Calls this page
                    End If
                End If
            Next
            If Incorrect Then
                MsgBox("Username, Password are Incorrect, Please Try Again.")                   'Sends the error mesasage to the user if the password or username is incorrect
            End If
        Catch ex As Exception
            MsgBox(ex.Message)                  'Sends an error message to the user if a error occurs
        End Try
        txtPassword.Clear()                     'Clears the value in  the password textbox
    End Sub
    Sub GetData()
        Dim Con As New SqlConnection(ConString)
        Dim Cmd As String = "SELECT Username, Password FROM Student"
        Dim Command As New SqlCommand
        Dim reader As SqlDataReader
        i = 0           'Assigns the value 0 to i
        Try
            If Con.State = ConnectionState.Closed Then              'Checks if the database connection state is closed
                Con.Open()              'Sets the connection state to open
            End If
            Command = Con.CreateCommand                 'Creates the settings for a new sql command
            Command.CommandText = "SELECT COUNT(*) FROM Student"            'Sets the command text
            TotalRows = CInt(Command.ExecuteScalar)                 'Executes the command query and stores the integer converted result in the variable
            ReDim User(TotalRows)
            Command = Con.CreateCommand             'Creates the settings for a new sql command 
            Command.CommandText = Cmd               'Sets the command text
            reader = Command.ExecuteReader          'Executes the command query and stores the results in the sql data reader
            Do While reader.Read                    'Checks if there is more values to be read
                User(i).Username = reader.GetValue(0)           'Assigns the value inside the reader to the variable
                User(i).Password = reader.GetValue(1)           'Assigns the value inside the reader to the variable
                i += 1              'Increments i by 1
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)                 'Sends an error message to the user if a error occurs
        Finally
            Con.Close()                 'Sets the connection state to closed
        End Try
    End Sub
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        AdminLogin.Show()           'Opens the AdminLogin Page
        Me.Close()                  'Closes this page
    End Sub
    Function Encrypted(ByRef password As String)            'See AddData Page, Encryption Procedure
        Dim word As String = password
        Dim Encryption As String
        Dim recursion As Integer
        If password <> Nothing Then
            word = AddWord(word)
            Encryption = railway(word, recursion)
        Else
            Encryption = Nothing
        End If
        Return Encryption
    End Function
    Function AddWord(ByRef word As String)              'See AddData Page, AddWord Procedure
        Dim Letter As Char
        Do Until word.Length >= 30
            For i = 1 To word.Length
                Letter = Mid(word, i, 1)
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
    End Function
    Function railway(ByRef word As String, ByRef Recursion As Integer)              'See AddData Page, railway Procedure
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
    End Function
    Private Sub txtPassword_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1_Click(sender, e)                'Handles if the enter key is pressed to enter details to login
        End If
    End Sub
End Class
