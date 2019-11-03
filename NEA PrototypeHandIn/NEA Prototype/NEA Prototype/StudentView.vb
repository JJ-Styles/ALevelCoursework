Imports System.Data.Sql
Imports System.Data.SqlClient
Public Class StudentView
    Dim User() As Data
    Dim User1 As Data
    Structure Data
        Dim Firstname, Surname, Qualification, BSCID, result As String
        Dim ResultP, ResultS, ResultW, ResultI, ResultOfPreviousAttemptP, ResultOfPreviousAttemptS, ResultOfPreviousAttemptW, ResultOfPreviousAttemptI As Decimal
        Dim TimeSinceLastAttempt As Date
        Dim ResultOfAttempt1, ResultOfAttempt2, ResultOfAttempt3 As Decimal
        Dim Attempts As Integer
    End Structure
    Private Sub StudentView_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txResultI2.Enabled = False              'Makes the Textboxes uneditable so the user cannot add data to them
        txResultI3.Enabled = False              'Makes the Textboxes uneditable so the user cannot add data to them
        txResultP2.Enabled = False              'Makes the Textboxes uneditable so the user cannot add data to them
        txResultP3.Enabled = False              'Makes the Textboxes uneditable so the user cannot add data to them
        txResultS2.Enabled = False              'Makes the Textboxes uneditable so the user cannot add data to them
        txResultS3.Enabled = False              'Makes the Textboxes uneditable so the user cannot add data to them
        txResultW2.Enabled = False              'Makes the Textboxes uneditable so the user cannot add data to them
        txResultW3.Enabled = False              'Makes the Textboxes uneditable so the user cannot add data to them
        GetStudentName()                    'Calls the procedure GetStudentName
        GetData()                    'Calls the procedure GetData
        TimeDifference()                    'Calls the procedureTimeDiffernce
        DisplayData()                    'Calls the procedure DisplayData
        PassFail()                    'Calls the procedure PassFail
    End Sub
    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        StudentLogin.Show()         'Opens the StudentLogin Page
        Me.Close()          'Closes this page
    End Sub
    Sub GetData()
        Dim i As Integer
        Dim totalrows As Integer
        Dim Con As New SqlConnection(ConString)
        Dim Cmd As String = "SELECT Student.ResultP, Student.ResultS, Student.ResultW, Student.ResultI FROM Student WHERE BSCID = '" & User1.BSCID & "'"
        Dim Command As New SqlCommand
        Dim reader As SqlDataReader
        Try
            If Con.State = ConnectionState.Closed Then                  'Checks if the connectiopn state is closed
                Con.Open()          'Sets the connection state to closed
            End If
            Command = Con.CreateCommand                 'Creates the settings for a new sql command
            Command.CommandText = "SELECT COUNT(*) FROM Test"               'Sets the command text
            totalrows = CInt(Command.ExecuteScalar)                 'Executes the command query and stores the integer converted result to the variable
            ReDim User(totalrows)
            Command = Con.CreateCommand             'Creates the settings for a new sql command
            Command.CommandText = Cmd           'Sets the command text
            reader = Command.ExecuteReader          'Executes the command query and stores the results in the sql data reader
            Do While reader.Read                'Checks if there is more values to be read
                If IsDBNull(reader.GetValue(0)) Then            'Checks if the valuen inside the reader is a database null
                    User1.ResultOfPreviousAttemptP = 0          'Assigns the value 0 to the variable if the value is a database null 
                Else
                    User1.ResultOfPreviousAttemptP = reader.GetValue(0)     'Assigns the value inside the reaer to the variable
                End If
                If IsDBNull(reader.GetValue(1)) Then            'Checks if the valuen inside the reader is a database null
                    User1.ResultOfPreviousAttemptS = 0          'Assigns the value 0 to the variable if the value is a database null
                Else
                    User1.ResultOfPreviousAttemptS = reader.GetValue(1)     'Assigns the value inside the reaer to the variable
                End If
                If IsDBNull(reader.GetValue(2)) Then            'Checks if the valuen inside the reader is a database null
                    User1.ResultOfPreviousAttemptW = 0          'Assigns the value 0 to the variable if the value is a database null
                Else
                    User1.ResultOfPreviousAttemptW = reader.GetValue(2)     'Assigns the value inside the reaer to the variable
                End If
                If IsDBNull(reader.GetValue(3)) Then            'Checks if the valuen inside the reader is a database null
                    User1.ResultOfPreviousAttemptI = 0          'Assigns the value 0 to the variable if the value is a database null
                Else
                    User1.ResultOfPreviousAttemptI = reader.GetValue(3)     'Assigns the value inside the reaer to the variable
                End If
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)         'Sends an error message to the user if a error occurs
        Finally
            Con.Close()                 'Sets the connection state to closed
        End Try
        Try
            If Con.State = ConnectionState.Closed Then          'Checks if the connection state is closed
                Con.Open()              'Sets the connection to open
            End If
            Command = Con.CreateCommand             'Creates the settings for a new sql command 
            Command.CommandText = "SELECT TimeSinceLastResult FROM TEST WHERE BSCID = '" & User1.BSCID & "'"            'Sets th command text
            reader = Command.ExecuteReader                  'Executes the command query and stores the results in the sql data reader
            Do While reader.Read                    'Checks if the reader has more values to be read
                If IsDBNull(reader.GetValue(0)) Then            'Checks if the valuen inside the reader is a database null
                    User(i).TimeSinceLastAttempt = Nothing          'Assigns the nothing value to the variable if the value in the reader is a database null
                Else
                    User(i).TimeSinceLastAttempt = reader.GetValue(0)     'Assigns the value inside the reaer to the variable
                End If
                i += 1              'Increments i by 1
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)             'Sends an error message to the user if a error occurs
        Finally
            Con.Close()                 'Sets the connection state to closed
        End Try
        i = 0               'Assigns the value 0 to i 
        Try
            If Con.State = ConnectionState.Closed Then          'Checks if the connection state is closed
                Con.Open()          'Sets the connection state to open
            End If
            Command = Con.CreateCommand     'Creates the settings for a new sql connection
            Command.CommandText = "SELECT ResultOfAttempt1, ResultOfAttempt2, ResultOfAttempt3 FROM TEST WHERE BSCID = '" & User1.BSCID & "'"           'Sets the command text
            reader = Command.ExecuteReader              'Executes the command query and stores the results in the sql data reader
            Do While reader.Read                'Checks if the reader has more values to be read
                If IsDBNull(reader.GetValue(0)) Then            'Checks if the valuen inside the reader is a database null
                    User(i).ResultOfAttempt1 = 0          'Assigns the value 0 to the variable if the value in the reader is a database null
                Else
                    User(i).ResultOfAttempt1 = reader.GetValue(0)     'Assigns the value inside the reaer to the variable
                End If
                If IsDBNull(reader.GetValue(1)) Then            'Checks if the valuen inside the reader is a database null
                    User(i).ResultOfAttempt2 = 0          'Assigns the value 0 to the variable if the value in the reader is a database null
                Else
                    User(i).ResultOfAttempt2 = reader.GetValue(1)     'Assigns the value inside the reaer to the variable
                End If
                If IsDBNull(reader.GetValue(2)) Then            'Checks if the valuen inside the reader is a database null
                    User(i).ResultOfAttempt3 = 0          'Assigns the value 0 to the variable if the value in the reader is a database null
                Else
                    User(i).ResultOfAttempt3 = reader.GetValue(2)     'Assigns the value inside the reaer to the variable
                End If
                i += 1                  'Increments i by 1
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)             'Sends an error message to the user if a error occurs
        Finally
            Con.Close()                 'Sets the connection state to open
        End Try
        i = 0               'Assigns the value 0 to i
        Try
            If Con.State = ConnectionState.Closed Then              'Checks if the connection state is set to closed
                Con.Open()              'Sets the connection state to open 
            End If
            Command = Con.CreateCommand     'Creates the settings for a new sql command
            Command.CommandText = "SELECT Attempts FROM TEST WHERE BSCID = '" & User1.BSCID & "'"       'Sets the command text
            reader = Command.ExecuteReader          'Executes the command query and stores the results in the sql data reader
            Do While reader.Read            ''Checks if the reader has more values to be read
                If IsDBNull(reader.GetValue(0)) Then            'Checks if the valuen inside the reader is a database null
                    User(i).Attempts = 0          'Assigns the value 0 to the variable if the value in the reader is a database null
                Else
                    User(i).Attempts = reader.GetValue(0)     'Assigns the value inside the reaer to the variable
                End If
                i += 1          'Increments i by 1
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)             'Sends an error message to the user if a error occurs
        Finally
            Con.Close()         'Sets the connection state to closed
        End Try
    End Sub
    Sub GetStudentName()
        Dim Con As New SqlConnection(ConString)
        Dim Cmd As String = "SELECT Firstname, Surname, BSCID FROM Student WHERE Username = '" & StudentUser & "'"
        Dim Command As New SqlCommand
        Dim reader As SqlDataReader
        Try
            If Con.State = ConnectionState.Closed Then          'Checks if the connection state is closed
                Con.Open()      'Sets the connection state to closed
            End If
            Command = Con.CreateCommand         'Creates the settings for a new sql command
            Command.CommandText = Cmd           'Sets the command text
            reader = Command.ExecuteReader          'Executes the command =query and stores the results to the sql data reader
            Do While reader.Read                'Checks if the reader has more values to be read
                User1.Firstname = reader.GetValue(0)            'Assigns the value inside the reader to the variable
                User1.Surname = reader.GetValue(1)            'Assigns the value inside the reader to the variable
                User1.BSCID = reader.GetValue(2)            'Assigns the value inside the reader to the variable
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)     'Sends an error message to the user if a error message 
        Finally
            Con.Close()         'Sets the connection state to closed
        End Try
        Try
            tbxName.Text = User1.Firstname & " " & User1.Surname        'Assigns the fullname of the student from the variables to the textbox
        Catch ex As Exception
            MsgBox(ex.Message)      'Sends an error message to the user an error occurs
        End Try
    End Sub
    Sub AddToTable(ByVal Cmd As String)
        Dim Con As New SqlConnection(ConString)
        Dim Command As New SqlCommand(Cmd, Con)
        Dim Writer As New SqlDataAdapter
        Try
            If Con.State = ConnectionState.Closed Then                  'Checks if the connection state is set to closed
                Con.Open()          'Sets the connection state to open
            ElseIf Con.State <> ConnectionState.Closed Then                  'Checks if the connection state is not set to closed
                Con.Close()          'Sets the connection state to closed
                Con.Open()          'Sets the connection state to open
            End If
            Command = Con.CreateCommand         'Creates the settings for a new sql command
            Command.CommandText = Cmd           'Sets the command text
            Command.ExecuteNonQuery()           'Executes the command query
        Catch ex As Exception
            MsgBox(ex.Message)          'Sends an error message to the user if a error occurs
        Finally
            Con.Close()         'Sets the connection state to closed
        End Try
    End Sub
    Private Sub txResultP1_TextChanged(sender As Object, e As EventArgs) Handles txResultP1.TextChanged
        Dim Cmd As String
        If txResultP1.TextLength >= 2 And txResultP1.TextLength <= 3 Then       'Checks if the variable length is between 2 and 3 
            If txResultP1.Text >= 0 And txResultP1.Text <= 100 Then         'Checks if the textbox value is between 0 and 100
                User(0).TimeSinceLastAttempt = Date.Now.ToLocalTime         'Stores the current system in the variable
                User(0).ResultP = txResultP1.Text           'Stores the value from the textbox to the variable
                User(0).Attempts += 1               'Increments the variable by 1
                txResultP1.Enabled = False          'Makes the textbox uneditable by the user
                Cmd = "UPDATE Test SET TimeSinceLastResult = '" & User(0).TimeSinceLastAttempt & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'PR2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET ResultOfAttempt1 = '" & User(0).ResultP & "' WHERE  BSCID = '" & User1.BSCID & "' AND UnitCode = 'PR2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET Attempts = '" & User(0).Attempts & "' WHERE BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                If User(0).ResultP > User1.ResultOfPreviousAttemptP Then                'Checks if the current entered result is greater than the highest stored value
                    Cmd = "UPDATE Student SET ResultP = '" & User(0).ResultP & "' WHERE BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                    AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                    User1.ResultOfPreviousAttemptP = User(0).ResultP            'Assigns the current value into the highest result variable
                    DisplayData()       'Calls the procedure DisplayData
                    SetBoundaries()     'Calls the procedure SetBoundries
                    PassFail()          'Calls the procedure PassFail
                End If
            Else
                MsgBox("Value Must be between 0 and 100." & vbCrLf & "No Percentage Sign Required.")            'Sends an error message to the user if a value between 0 and 100 isnt entered
            End If
        End If
    End Sub
    Private Sub txResultP2_TextChanged(sender As Object, e As EventArgs) Handles txResultP2.TextChanged
        Dim Cmd As String
        If txResultP2.TextLength >= 2 And txResultP2.TextLength <= 3 Then       'Checks if the variable length is between 2 and 3 
            If txResultP2.Text >= 0 And txResultP2.Text <= 100 Then         'Checks if the textbox value is between 0 and 100
                User(0).TimeSinceLastAttempt = Date.Now         'Stores the current system in the variable
                User(1).ResultP = txResultP2.Text           'Stores the value from the textbox to the variable
                User(0).Attempts += 1               'Increments the variable by 1
                txResultP2.Enabled = False          'Makes the textbox uneditable by the user
                Cmd = "UPDATE Test SET TimeSinceLastResult = '" & User(0).TimeSinceLastAttempt & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'PR2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET ResultOfAttempt2 = '" & User(1).ResultP & "' WHERE BSCID ='" & User1.BSCID & "' AND UnitCode = 'PR2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET Attempts = '" & User(0).Attempts & "' WHERE BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                If User(1).ResultP > User1.ResultOfPreviousAttemptP Then                'Checks if the current entered result is greater than the highest stored value
                    Cmd = "UPDATE Student SET Student.ResultP = '" & User(1).ResultP & "' WHERE Student.BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                    AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                    User1.ResultOfPreviousAttemptP = User(1).ResultP            'Assigns the current value into the highest result variable
                    DisplayData()       'Calls the procedure DisplayData
                    SetBoundaries()     'Calls the procedure SetBoundries
                    PassFail()          'Calls the procedure PassFail
                End If
            Else
                MsgBox("Value Must be between 0 and 100." & vbCrLf & "No Percentage Sign Required.")            'Sends an error message to the user if a value between 0 and 100 isnt entered
            End If
        End If
    End Sub
    Private Sub txResultP3_TextChanged(sender As Object, e As EventArgs) Handles txResultP3.TextChanged
        Dim Cmd As String
        If txResultP3.TextLength >= 2 And txResultP3.TextLength <= 3 Then       'Checks if the variable length is between 2 and 3 
            If txResultP3.Text >= 0 And txResultP3.Text <= 100 Then         'Checks if the textbox value is between 0 and 100
                User(0).TimeSinceLastAttempt = Date.Now         'Stores the current system in the variable
                User(2).ResultP = txResultP3.Text           'Stores the value from the textbox to the variable
                User(0).Attempts += 1               'Increments the variable by 1
                txResultP3.Enabled = False          'Makes the textbox uneditable by the user
                Cmd = "UPDATE Test SET TimeSinceLastResult = '" & User(0).TimeSinceLastAttempt & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'PR2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET ResultOfAttempt3 = '" & User(2).ResultP & "'  WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'PR2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET Attempts = '" & User(0).Attempts & "' WHERE BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                If User(2).ResultP > User1.ResultOfPreviousAttemptP Then                'Checks if the current entered result is greater than the highest stored value
                    Cmd = "UPDATE Student SET Student.ResultP = '" & User(2).ResultP & "' WHERE Student.BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                    AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                    User1.ResultOfPreviousAttemptP = User(2).ResultP            'Assigns the current value into the highest result variable
                    DisplayData()       'Calls the procedure DisplayData
                    SetBoundaries()     'Calls the procedure SetBoundries
                    PassFail()          'Calls the procedure PassFail
                End If
            Else
                MsgBox("Value Must be between 0 and 100." & vbCrLf & "No Percentage Sign Required.")            'Sends an error message to the user if a value between 0 and 100 isnt entered
            End If
        End If
    End Sub
    Private Sub txResultS1_TextChanged(sender As Object, e As EventArgs) Handles txResultS1.TextChanged
        Dim Cmd As String
        If txResultS1.TextLength >= 2 And txResultS1.TextLength <= 3 Then       'Checks if the variable length is between 2 and 3 
            If txResultS1.Text >= 0 And txResultS1.Text <= 100 Then         'Checks if the textbox value is between 0 and 100
                User(1).TimeSinceLastAttempt = Date.Now         'Stores the current system in the variable
                User(0).ResultS = txResultS1.Text           'Stores the value from the textbox to the variable
                User(1).Attempts += 1               'Increments the variable by 1
                txResultS1.Enabled = False          'Makes the textbox uneditable by the user
                Cmd = "UPDATE Test SET TimeSinceLastResult = '" & User(1).TimeSinceLastAttempt & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SS2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET ResultOfAttempt1 = '" & User(0).ResultS & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SS2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET Attempts = '" & User(1).Attempts & "' WHERE BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                If User(0).ResultS > User1.ResultOfPreviousAttemptS Then                'Checks if the current entered result is greater than the highest stored value
                    Cmd = "UPDATE Student SET Student.ResultS = '" & User(0).ResultS & "' WHERE Student.BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                    AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                    User1.ResultOfPreviousAttemptS = User(0).ResultS            'Assigns the current value into the highest result variable
                    DisplayData()       'Calls the procedure DisplayData
                    SetBoundaries()     'Calls the procedure SetBoundries
                    PassFail()          'Calls the procedure PassFail
                End If
            Else
                MsgBox("Value Must be between 0 and 100." & vbCrLf & "No Percentage Sign Required.")            'Sends an error message to the user if a value between 0 and 100 isnt entered
            End If
        End If
    End Sub
    Private Sub txResultS2_TextChanged(sender As Object, e As EventArgs) Handles txResultS2.TextChanged
        Dim Cmd As String
        If txResultS2.TextLength >= 2 And txResultS2.TextLength <= 3 Then       'Checks if the variable length is between 2 and 3 
            If txResultS2.Text >= 0 And txResultS2.Text <= 100 Then         'Checks if the textbox value is between 0 and 100
                User(1).TimeSinceLastAttempt = Date.Now         'Stores the current system in the variable
                User(1).ResultS = txResultS2.Text           'Stores the value from the textbox to the variable
                User(1).Attempts += 1               'Increments the variable by 1
                txResultS2.Enabled = False          'Makes the textbox uneditable by the user
                Cmd = "UPDATE Test SET TimeSinceLastResult = '" & User(1).TimeSinceLastAttempt & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SS2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET ResultOfAttempt2 = '" & User(1).ResultS & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SS2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET Attempts = '" & User(1).Attempts & "' WHERE BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                If User(1).ResultS > User1.ResultOfPreviousAttemptS Then                'Checks if the current entered result is greater than the highest stored value
                    Cmd = "UPDATE Student SET Student.ResultS = '" & User(1).ResultS & "' WHERE Student.BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                    AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                    User1.ResultOfPreviousAttemptS = User(1).ResultS            'Assigns the current value into the highest result variable
                    DisplayData()       'Calls the procedure DisplayData
                    SetBoundaries()     'Calls the procedure SetBoundries
                    PassFail()          'Calls the procedure PassFail
                End If
            Else
                MsgBox("Value Must be between 0 and 100." & vbCrLf & "No Percentage Sign Required.")            'Sends an error message to the user if a value between 0 and 100 isnt entered
            End If
        End If
    End Sub
    Private Sub txResultS3_TextChanged(sender As Object, e As EventArgs) Handles txResultS3.TextChanged
        Dim Cmd As String
        If txResultS3.TextLength >= 2 And txResultS3.TextLength <= 3 Then       'Checks if the variable length is between 2 and 3 
            If txResultS3.Text >= 0 And txResultS3.Text <= 100 Then         'Checks if the textbox value is between 0 and 100
                User(1).TimeSinceLastAttempt = Date.Now         'Stores the current system in the variable
                User(2).ResultS = txResultS3.Text           'Stores the value from the textbox to the variable
                User(1).Attempts += 1               'Increments the variable by 1
                txResultS3.Enabled = False          'Makes the textbox uneditable by the user
                Cmd = "UPDATE Test SET TimeSinceLastResult = '" & User(1).TimeSinceLastAttempt & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SS2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET ResultOfAttempt3 = '" & User(2).ResultS & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SS2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET Attempts = '" & User(1).Attempts & "' WHERE BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                If User(2).ResultS > User1.ResultOfPreviousAttemptS Then                'Checks if the current entered result is greater than the highest stored value
                    Cmd = "UPDATE Student SET Student.ResultS = '" & User(2).ResultS & "' WHERE Student.BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                    AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                    User1.ResultOfPreviousAttemptS = User(2).ResultS            'Assigns the current value into the highest result variable
                    DisplayData()       'Calls the procedure DisplayData
                    SetBoundaries()     'Calls the procedure SetBoundries
                    PassFail()          'Calls the procedure PassFail
                End If
            Else
                MsgBox("Value Must be between 0 and 100." & vbCrLf & "No Percentage Sign Required.")            'Sends an error message to the user if a value between 0 and 100 isnt entered
            End If
        End If
    End Sub
    Private Sub txResultW1_TextChanged(sender As Object, e As EventArgs) Handles txResultW1.TextChanged
        Dim Cmd As String
        If txResultW1.TextLength >= 2 And txResultW1.TextLength <= 3 Then       'Checks if the variable length is between 2 and 3 
            If txResultW1.Text >= 0 And txResultW1.Text <= 100 Then          'Checks if the textbox value is between 0 and 100
                User(2).TimeSinceLastAttempt = Date.Now         'Stores the current system in the variable
                User(0).ResultW = txResultW1.Text           'Stores the value from the textbox to the variable
                User(2).Attempts += 1               'Increments the variable by 1
                txResultW1.Enabled = False          'Makes the textbox uneditable by the user
                Cmd = "UPDATE Test SET TimeSinceLastResult = '" & User(2).TimeSinceLastAttempt & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'WP2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET ResultOfAttempt1 = '" & User(0).ResultW & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'WP2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET Attempts = '" & User(2).Attempts & "' WHERE BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                If User(0).ResultW > User1.ResultOfPreviousAttemptW Then                'Checks if the current entered result is greater than the highest stored value
                    Cmd = "UPDATE Student SET Student.ResultW = '" & User(0).ResultW & "' WHERE Student.BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                    AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                    User1.ResultOfPreviousAttemptW = User(0).ResultW            'Assigns the current value into the highest result variable
                    DisplayData()       'Calls the procedure DisplayData
                    SetBoundaries()     'Calls the procedure SetBoundries
                    PassFail()          'Calls the procedure PassFail
                End If
            Else
                MsgBox("Value Must be between 0 and 100." & vbCrLf & "No Percentage Sign Required.")            'Sends an error message to the user if a value between 0 and 100 isnt entered
            End If
        End If
    End Sub
    Private Sub txResultW2_TextChanged(sender As Object, e As EventArgs) Handles txResultW2.TextChanged
        Dim Cmd As String
        If txResultW2.TextLength >= 2 And txResultW2.TextLength <= 3 Then       'Checks if the variable length is between 2 and 3 
            If txResultW2.Text >= 0 And txResultW2.Text <= 100 Then         'Checks if the textbox value is between 0 and 100
                User(2).TimeSinceLastAttempt = Date.Now         'Stores the current system in the variable
                User(1).ResultW = txResultW2.Text           'Stores the value from the textbox to the variable
                User(2).Attempts += 1               'Increments the variable by 1
                txResultW2.Enabled = False          'Makes the textbox uneditable by the user
                Cmd = "UPDATE Test SET TimeSinceLastResult = '" & User(2).TimeSinceLastAttempt & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'WP2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET ResultOfAttempt2 = '" & User(1).ResultW & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'WP2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET Attempts = '" & User(2).Attempts & "' WHERE BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                If User(0).ResultW > User1.ResultOfPreviousAttemptW Then                'Checks if the current entered result is greater than the highest stored value
                    Cmd = "UPDATE Student SET Student.ResultW = '" & User(1).ResultW & "' WHERE Student.BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                    AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                    User1.ResultOfPreviousAttemptW = User(1).ResultW            'Assigns the current value into the highest result variable
                    DisplayData()       'Calls the procedure DisplayData
                    SetBoundaries()     'Calls the procedure SetBoundries
                    PassFail()          'Calls the procedure PassFail
                End If
            Else
                MsgBox("Value Must be between 0 and 100." & vbCrLf & "No Percentage Sign Required.")            'Sends an error message to the user if a value between 0 and 100 isnt entered
            End If
        End If
    End Sub
    Private Sub txResultW3_TextChanged(sender As Object, e As EventArgs) Handles txResultW3.TextChanged
        Dim Cmd As String
        If txResultW3.TextLength >= 2 And txResultW3.TextLength <= 3 Then       'Checks if the variable length is between 2 and 3 
            If txResultW3.Text >= 0 And txResultW3.Text <= 100 Then         'Checks if the textbox value is between 0 and 100
                User(2).TimeSinceLastAttempt = Date.Now         'Stores the current system in the variable
                User(2).ResultW = txResultW3.Text           'Stores the value from the textbox to the variable
                User(2).Attempts += 1               'Increments the variable by 1
                txResultW3.Enabled = False          'Makes the textbox uneditable by the user
                Cmd = "UPDATE Test SET TimeSinceLastResult = '" & User(2).TimeSinceLastAttempt & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'WP2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET ResultOfAttempt3 = '" & User(2).ResultW & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'WP2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET Attempts = '" & User(2).Attempts & "' WHERE BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                If User(2).ResultW > User1.ResultOfPreviousAttemptW Then                'Checks if the current entered result is greater than the highest stored value
                    Cmd = "UPDATE Student SET Student.ResultW = '" & User(2).ResultW & "' WHERE Student.BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                    AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                    User1.ResultOfPreviousAttemptW = User(2).ResultW            'Assigns the current value into the highest result variable
                    DisplayData()       'Calls the procedure DisplayData
                    SetBoundaries()     'Calls the procedure SetBoundries
                    PassFail()          'Calls the procedure PassFail
                End If
            Else
                MsgBox("Value Must be between 0 and 100." & vbCrLf & "No Percentage Sign Required.")            'Sends an error message to the user if a value between 0 and 100 isnt entered
            End If
        End If
    End Sub
    Private Sub txResultI1_TextChanged(sender As Object, e As EventArgs) Handles txResultI1.TextChanged
        Dim Cmd As String
        If txResultI1.TextLength >= 2 And txResultI1.TextLength <= 3 Then       'Checks if the variable length is between 2 and 3 
            If txResultI1.Text >= 0 And txResultI1.Text <= 100 Then         'Checks if the textbox value is between 0 and 100
                User(3).TimeSinceLastAttempt = Date.Now         'Stores the current system in the variable
                User(0).ResultI = txResultI1.Text           'Stores the value from the textbox to the variable
                User(3).Attempts += 1               'Increments the variable by 1
                txResultI1.Enabled = False          'Makes the textbox uneditable by the user
                Cmd = "UPDATE Test SET TimeSinceLastResult = '" & User(3).TimeSinceLastAttempt & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SIP2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET ResultOfAttempt1 = '" & User(0).ResultI & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SIP2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET Attempts = '" & User(3).Attempts & "' WHERE BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                If User(0).ResultI > User1.ResultOfPreviousAttemptI Then                'Checks if the current entered result is greater than the highest stored value
                    Cmd = "UPDATE Student SET Student.ResultI = '" & User(0).ResultI & "' WHERE Student.BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                    AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                    User1.ResultOfPreviousAttemptI = User(0).ResultI            'Assigns the current value into the highest result variable
                    DisplayData()       'Calls the procedure DisplayData
                    SetBoundaries()     'Calls the procedure SetBoundries
                    PassFail()          'Calls the procedure PassFail
                End If
            Else
                MsgBox("Value Must be between 0 and 100." & vbCrLf & "No Percentage Sign Required.")            'Sends an error message to the user if a value between 0 and 100 isnt entered
            End If
        End If
    End Sub
    Private Sub txResultI2_TextChanged(sender As Object, e As EventArgs) Handles txResultI2.TextChanged
        Dim Cmd As String
        If txResultI2.TextLength >= 2 And txResultI2.TextLength <= 3 Then       'Checks if the variable length is between 2 and 3 
            If txResultI2.Text >= 0 And txResultI2.Text <= 100 Then         'Checks if the textbox value is between 0 and 100
                User(3).TimeSinceLastAttempt = Date.Now         'Stores the current system in the variable
                User(1).ResultI = txResultI2.Text           'Stores the value from the textbox to the variable
                User(3).Attempts += 1               'Increments the variable by 1
                txResultI2.Enabled = False          'Makes the textbox uneditable by the user
                Cmd = "UPDATE Test SET TimeSinceLastResult = '" & User(3).TimeSinceLastAttempt & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SIP2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET ResultOfAttempt2 = '" & User(1).ResultI & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SIP2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET Attempts = '" & User(3).Attempts & "' WHERE BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                If User(1).ResultI > User1.ResultOfPreviousAttemptI Then                'Checks if the current entered result is greater than the highest stored value
                    Cmd = "UPDATE Student SET Student.ResultW = '" & User(1).ResultI & "' WHERE Student.BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                    AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                    User1.ResultOfPreviousAttemptI = User(1).ResultI            'Assigns the current value into the highest result variable
                    DisplayData()       'Calls the procedure DisplayData
                    SetBoundaries()     'Calls the procedure SetBoundries
                    PassFail()          'Calls the procedure PassFail
                End If
            Else
                MsgBox("Value Must be between 0 and 100." & vbCrLf & "No Percentage Sign Required.")            'Sends an error message to the user if a value between 0 and 100 isnt entered
            End If
        End If
    End Sub
    Private Sub txResultI3_TextChanged(sender As Object, e As EventArgs) Handles txResultI3.TextChanged
        Dim Cmd As String
        If txResultI3.TextLength >= 2 And txResultI3.TextLength <= 3 Then       'Checks if the variable length is between 2 and 3 
            If txResultI3.Text >= 0 And txResultI3.Text <= 100 Then         'Checks if the textbox value is between 0 and 100 
                User(3).TimeSinceLastAttempt = Date.Now         'Stores the current system in the variable
                User(2).ResultI = txResultI3.Text           'Stores the value from the textbox to the variable 
                User(3).Attempts += 1                'Increments the variable by 1
                txResultI3.Enabled = False          'Makes the textbox uneditable by the user
                Cmd = "UPDATE Test SET TimeSinceLastResult = '" & User(3).TimeSinceLastAttempt & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SIP2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET ResultOfAttempt3 = '" & User(2).ResultI & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SIP2'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                Cmd = "UPDATE Test SET Attempts = '" & User(3).Attempts & "' WHERE BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                If User(2).ResultI > User1.ResultOfPreviousAttemptI Then                'Checks if the current entered result is greater than the highest stored value
                    Cmd = "UPDATE Student SET Student.ResultW = '" & User(2).ResultI & "' WHERE Student.BSCID = '" & User1.BSCID & "'"           'Assigns the string value to the variable
                    AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
                    User1.ResultOfPreviousAttemptI = User(2).ResultI            'Assigns the current value into the highest result variable
                    DisplayData()       'Calls the procedure DisplayData
                    SetBoundaries()     'Calls the procedure SetBoundries
                    PassFail()          'Calls the procedure PassFail
                End If
            Else
                MsgBox("Value Must be between 0 and 100." & vbCrLf & "No Percentage Sign Required.")            'Sends an error message to the user if a value between 0 and 100 isnt entered
            End If
        End If
    End Sub
    Sub TimeDifference()
        Dim Time(3) As TimeSpan
        Dim TimeSinceLastAttempt(3) As Double
        For i = 0 To 3
            If User(i).TimeSinceLastAttempt = Nothing Then  'Checks if the variable contains a null value
                TimeSinceLastAttempt(i) = 0         'Assigns the value 0 to the variable
            Else
                Time(i) = Date.Now.Subtract(User(i).TimeSinceLastAttempt.ToLocalTime)        'Calculates the Time Difference between the time of the system now and the variable
                TimeSinceLastAttempt(i) = Math.Floor(CDbl(Time(i).TotalHours))               'Formats the time scale into a writable format
            End If
        Next
        If TimeSinceLastAttempt(0) >= 48 And User(0).ResultOfAttempt1 <> Nothing Then           'Checks if 48 hours has passed since the last attempt was entered
            txResultP2.Enabled = True                   'Makes the textbox editable by the user
        End If
        If TimeSinceLastAttempt(0) >= 48 And User(0).ResultOfAttempt2 <> Nothing Then           'Checks if 48 hours has passed since the last attempt was entered
            txResultP3.Enabled = True                                           'Makes the textbox editable by the user
            If User(0).ResultOfAttempt3 <> 0 Then
                txResultP3.Enabled = False
            End If
        End If
        If TimeSinceLastAttempt(1) >= 48 And User(1).ResultOfAttempt1 <> Nothing Then           'Checks if 48 hours has passed since the last attempt was entered
            txResultS2.Enabled = True                   'Makes the textbox editable by the user
        End If
        If TimeSinceLastAttempt(1) >= 48 And User(1).ResultOfAttempt2 <> Nothing Then           'Checks if 48 hours has passed since the last attempt was entered
            txResultS3.Enabled = True                   'Makes the textbox editable by the user
            If User(0).ResultOfAttempt3 <> 0 Then
                txResultS3.Enabled = False
            End If

        End If
        If TimeSinceLastAttempt(2) >= 48 And User(2).ResultOfAttempt1 <> Nothing Then           'Checks if 48 hours has passed since the last attempt was entered
            txResultW2.Enabled = True                   'Makes the textbox editable by the user
        End If
        If TimeSinceLastAttempt(2) >= 48 And User(2).ResultOfAttempt2 <> Nothing Then           'Checks if 48 hours has passed since the last attempt was entered
            txResultW3.Enabled = True                   'Makes the textbox editable by the user
            If User(0).ResultOfAttempt3 <> 0 Then
                txResultW3.Enabled = False
            End If

        End If
        If TimeSinceLastAttempt(3) >= 48 And User(3).ResultOfAttempt1 <> Nothing Then           'Checks if 48 hours has passed since the last attempt was entered
            txResultI2.Enabled = True                   'Makes the textbox editable by the user
        End If
        If TimeSinceLastAttempt(3) >= 48 And User(3).ResultOfAttempt2 <> Nothing Then           'Checks if 48 hours has passed since the last attempt was entered
            txResultI3.Enabled = True                   'Makes the textbox editable by the user
            If User(0).ResultOfAttempt3 <> 0 Then
                txResultI3.Enabled = False
            End If

        End If

        tbxTimeSinceLastAttemptP.Text = TimeSinceLastAttempt(0)         'Assigns the value inside the variable to the textbox
        tbxTimeSinceLastAttemptS.Text = TimeSinceLastAttempt(1)         'Assigns the value inside the variable to the textbox
        tbxTimeSinceLastAttemptW.Text = TimeSinceLastAttempt(2)         'Assigns the value inside the variable to the textbox
        tbxTimeSinceLastAttemptI.Text = TimeSinceLastAttempt(3)         'Assigns the value inside the variable to the textbox
    End Sub
    Sub DisplayData()
        tbxCurrentGradeP.Text = User1.ResultOfPreviousAttemptP         'Assigns the value inside the variable to the textbox
        tbxCurrentGradeS.Text = User1.ResultOfPreviousAttemptS         'Assigns the value inside the variable to the textbox
        tbxCurrentGradeW.Text = User1.ResultOfPreviousAttemptW         'Assigns the value inside the variable to the textbox
        tbxCurrentGradeI.Text = User1.ResultOfPreviousAttemptI         'Assigns the value inside the variable to the textbox
    End Sub
    Sub SetBoundaries()
        Dim Cmd As String
        Dim Valid As Boolean
        If User1.ResultOfPreviousAttemptP <> Nothing Then           'Checks to see if the variable isnt empty
            If User1.ResultOfPreviousAttemptS <> Nothing Then           'Checks to see if the variable isnt empty
                If User1.ResultOfPreviousAttemptW <> Nothing Then           'Checks to see if the variable isnt empty
                    If User1.ResultOfPreviousAttemptI <> Nothing Then           'Checks to see if the variable isnt empty
                        Valid = GradeOverPassRate()             'Calls the procedure GradeOverPassRate and stores the output in the variable
                        If Valid Then                   'Checks if the variable is equal to true
                            CalculateBoundries()        'Calls the procedure CalculateBoundries
                        Else
                            User1.Qualification = "Fail"     'Assigns the value Fail to the variable
                        End If
                    End If
                End If
            End If
        End If
        Cmd = "UPDATE Student SET Qualification = '" & User1.Qualification & "' WHERE BSCID = '" & User1.BSCID & "'"            'Assigns the string to the variable
        AddToTable(Cmd)         'Calls the procedure AddToTable and passes the variable Cmd into it
    End Sub
    Function GradeOverPassRate() As Boolean
        If User1.ResultOfPreviousAttemptP < 75 Then        'Checks if the Result is less than the pass rate
            Return False                'Returns the value false to the previous procedure
        End If
        If User1.ResultOfPreviousAttemptS < 75 Then        'Checks if the Result is less than the pass rate
            Return False                'Returns the value false to the previous procedure
        End If
        If User1.ResultOfPreviousAttemptW < 75 Then        'Checks if the Result is less than the pass rate
            Return False                'Returns the value false to the previous procedure
        End If
        If User1.ResultOfPreviousAttemptI < 55 Then        'Checks if the Result is less than the pass rate
            Return False                'Returns the value false to the previous procedure
        End If
        Return True                'Returns the value True to the previous procedure
    End Function
    Sub CalculateBoundries()
        Dim result As Double
        result = User1.ResultOfPreviousAttemptP + User1.ResultOfPreviousAttemptS + User1.ResultOfPreviousAttemptW + User1.ResultOfPreviousAttemptI          'Adds all of the highest results together
        result /= 4         'Divides the variable by 4 because there is 4 tests
        If result < 70 Then         'Checks if the variable is below 70
            User1.Qualification = "Fail"                 'Assigns the value fail to the variable
        ElseIf result >= 70 And result < 75 Then         'Checks if the variable is greater than or equal to 70 and less than 75
            User1.Qualification = "Pass"                 'Assigns the value Pass to the variable
        ElseIf result >= 75 And result < 80 Then         'Checks if the variable is greater than or equal to 75 and less than 80
            User1.Qualification = "Merit"                 'Assigns the value Merit to the variable
        ElseIf result >= 80 And result < 85 Then         'Checks if the variable is greater than or equal to 80 and less than 85
            User1.Qualification = "Destinction"                 'Assigns the value Destinction to the variable
        ElseIf result >= 85 And result <= 100 Then         'Checks if the variable is greater than or equal to 85 and less than 100
            User1.Qualification = "Distinction*"                 'Assigns the value Destinction* to the variable
        End If
    End Sub
    Sub PassFail()
        Dim Cmd As String = ""
        If User1.ResultOfPreviousAttemptP <> 0 Then             'Checks the result doesnt equal 0
            txResultP1.Enabled = False      'Makes the textboxes uneditable to the user
            If User1.ResultOfPreviousAttemptP >= 70 And User1.ResultOfPreviousAttemptP <= 100 Then      'Checks if the result is between 70 and 100
                User1.result = "Passed"             'Assigns the value passed to the variable
                Cmd = "UPDATE Test SET Result = '" & User1.result & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'PR2'"            'Assigns the string to the Variable
                AddToTable(Cmd)             'Calls the Procedure AddToTable and passes through the variable Cmd
            ElseIf User1.ResultOfPreviousAttemptP < 70 And User1.ResultOfPreviousAttemptP >= 0 Then
                User1.result = "Failed"             'Assigns the value Failed to the variable
                Cmd = "UPDATE Test SET Result = '" & User1.result & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'PR2'"            'Assigns the string to the Variable
                AddToTable(Cmd)             'Calls the Procedure AddToTable and passes through the variable Cmd
            End If
        End If
        If User1.ResultOfPreviousAttemptW <> 0 Then             'Checks the result doesnt equal 0
            txResultW1.Enabled = False      'Makes the textboxes uneditable to the user
            If User1.ResultOfPreviousAttemptW >= 70 And User1.ResultOfPreviousAttemptW <= 100 Then      'Checks if the result is between 70 and 100
                User1.result = "Passed"             'Assigns the value passed to the variable
                Cmd = "UPDATE Test SET Result = '" & User1.result & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'WP2'"            'Assigns the string to the Variable
                AddToTable(Cmd)             'Calls the Procedure AddToTable and passes through the variable Cmd
            ElseIf User1.ResultOfPreviousAttemptW < 70 And User1.ResultOfPreviousAttemptW >= 0 Then
                User1.result = "Failed"             'Assigns the value Failed to the variable
                Cmd = "UPDATE Test SET Result = '" & User1.result & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'WP2'"            'Assigns the string to the Variable
                AddToTable(Cmd)             'Calls the Procedure AddToTable and passes through the variable Cmd
            End If
        End If
        If User1.ResultOfPreviousAttemptS <> 0 Then             'Checks the result doesnt equal 0
            txResultS1.Enabled = False      'Makes the textboxes uneditable to the user
            If User1.ResultOfPreviousAttemptS >= 70 And User1.ResultOfPreviousAttemptS <= 100 Then      'Checks if the result is between 70 and 100
                User1.result = "Passed"             'Assigns the value passed to the variable
                Cmd = "UPDATE Test SET Result = '" & User1.result & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SS2'"            'Assigns the string to the Variable
                AddToTable(Cmd)             'Calls the Procedure AddToTable and passes through the variable Cmd
            ElseIf User1.ResultOfPreviousAttemptS < 70 And User1.ResultOfPreviousAttemptS >= 0 Then
                User1.result = "Failed"             'Assigns the value Failed to the variable
                Cmd = "UPDATE Test SET Result = '" & User1.result & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SS2'"            'Assigns the string to the Variable
                AddToTable(Cmd)             'Calls the Procedure AddToTable and passes through the variable Cmd
            End If
        End If
        If User1.ResultOfPreviousAttemptI <> 0 Then             'Checks the result doesnt equal 0
            txResultI1.Enabled = False      'Makes the textboxes uneditable to the user
            If User1.ResultOfPreviousAttemptI >= 50 And User1.ResultOfPreviousAttemptI <= 100 Then      'Checks if the result is between 70 and 100
                User1.result = "Passed"             'Assigns the value passed to the variable
                Cmd = "UPDATE Test SET Result = '" & User1.result & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SIP2'"            'Assigns the string to the Variable
                AddToTable(Cmd)             'Calls the Procedure AddToTable and passes through the variable Cmd
            ElseIf User1.ResultOfPreviousAttemptI < 50 And User1.ResultOfPreviousAttemptP >= 0 Then
                User1.result = "Failed"             'Assigns the value Failed to the variable
                Cmd = "UPDATE Test SET Result = '" & User1.result & "' WHERE BSCID = '" & User1.BSCID & "' AND UnitCode = 'SIP2'"            'Assigns the string to the Variable
                AddToTable(Cmd)             'Calls the Procedure AddToTable and passes through the variable Cmd
            End If
        End If
        If User(0).ResultOfAttempt2 <> 0 Then           'Checks if the result of attempt 2 doesnt equal 0
            txResultP2.Enabled = False                  'Enables the textbox so that it is user editable
        ElseIf User(0).ResultOfAttempt3 <> 0 Then           'Checks if the result of attempt 3 doesnt equal 0
            txResultP3.Enabled = False                  'Enables the textbox so that it is user editable
        End If
        If User(1).ResultOfAttempt2 <> 0 Then           'Checks if the result of attempt 2 doesnt equal 0
            txResultS2.Enabled = False                  'Enables the textbox so that it is user editable
        ElseIf User(1).ResultOfAttempt3 <> 0 Then           'Checks if the result of attempt 3 doesnt equal 0
            txResultS3.Enabled = False                  'Enables the textbox so that it is user editable
        End If
        If User(2).ResultOfAttempt2 <> 0 Then           'Checks if the result of attempt 2 doesnt equal 0
            txResultW2.Enabled = False                  'Enables the textbox so that it is user editable
        ElseIf User(2).ResultOfAttempt3 <> 0 Then           'Checks if the result of attempt 3 doesnt equal 0
            txResultW3.Enabled = False                  'Enables the textbox so that it is user editable
        End If
        If User(3).ResultOfAttempt2 <> 0 Then           'Checks if the result of attempt 2 doesnt equal 0
            txResultI2.Enabled = False                  'Enables the textbox so that it is user editable
        ElseIf User(3).ResultOfAttempt3 <> 0 Then           'Checks if the result of attempt 3 doesnt equal 0
            txResultI3.Enabled = False                  'Enables the textbox so that it is user editable
        End If
    End Sub
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("https://ecdluk.psionline.com/")
    End Sub
End Class