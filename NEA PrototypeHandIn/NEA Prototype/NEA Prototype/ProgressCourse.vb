Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class ProgressCourse

    Structure Data
        Dim ResultP, ResultS, ResultW, ResultI, passFailratio, AverageResult As Decimal
        Dim Invigilator, ClassName As String
        Dim NotTaken As Integer
    End Structure

    Structure Data1
        Dim result, ClassName As String
    End Structure

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        AdminView1.Show()       'Opens the AdminView page
        Me.Close()              'Closes this page
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        StudentLogin.Show()     'Opens the StudentLogin Page
        Me.Close()              'Closes this page
    End Sub

    Private Sub ProgressCourse_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetData()       'Calls the procedure GetData
    End Sub

    Sub GetData()
        Dim User() As Data
        Dim user1() As Data1
        Dim NoInClass(2) As Integer
        Dim i, n, x, y As Integer
        Dim TotalRows, NoOfRowsClass As Integer
        Dim ClassRequired() As String
        Dim Con As New SqlConnection(ConString)
        Dim Cmd As String = "SELECT Student.ResultP, Student.ResultS, Student.ResultW, Student.ResultI, Class.Invigilator, Class.ClassName FROM Student, Class WHERE Student.ClassName = Class.ClassName ORDER BY Class.ClassName"
        Dim Cmd1 As String = "SELECT Test.Result, Test.ClassName FROM Test, Student ORDER BY Student.ClassName"
        Dim Command As New SqlCommand
        Dim reader As SqlDataReader
        i = 0
        Try
            If Con.State = ConnectionState.Closed Then      'Checks if the connections is closed
                Con.Open()                      'Sets the connection to open
            Else
                Con.Close()             'Sets the connection to close
                Con.Open()              'Sets the connection to open
            End If
            Command = Con.CreateCommand         'Creates the settings for a new SQL Data Reader 
            Command.CommandText = "SELECT COUNT(*) FROM Class"      'Sets the command text
            NoOfRowsClass = CInt(Command.ExecuteScalar)             'Executes the command query and stores the integer converted result in the variable NoOfRowsClass
            ReDim ClassRequired(NoOfRowsClass - 1)
            Command = Con.CreateCommand         'Creates the settings for a new SQL Data Reader 
            Command.CommandText = "SELECT ClassName FROM Class"      'Sets the command text
            reader = Command.ExecuteReader                           'Executes the command query and stores the result in the SQL Data Reader
            Do While reader.Read                    'Checks if there is anymore values to be read
                ClassRequired(y) = CStr(reader.GetValue(0))     'Stores the value from the reader into the variable
                y += 1              'Increments the variable on each cycle of the iteration
            Loop
            Con.Close()             'Sets the connection to closed
            Con.Open()              'Sets theconnection to open
            y = 0                   'Assigns the value 0 to the variable
            Command = Con.CreateCommand         'Creates the settings for a new SQL Data Reader
            Command.CommandText = "SELECT COUNT(*) FROM student"      'Sets the command text
            TotalRows = CInt(Command.ExecuteScalar)             'Executes the command query and stores the integer converted result in the variable NoOfRowsClass
            ReDim User(TotalRows - 1)
            Command = Con.CreateCommand         'Creates the settings for a new SQL Data Reader
            For m = 0 To (ClassRequired.Length - 1)
                Command.CommandText = "SELECT Count(*) FROM Student WHERE ClassName = '" & ClassRequired(m) & "'"      'Sets the command text
                NoInClass(y) = CInt(Command.ExecuteScalar)             'Executes the command query and stores the integer converted result in the variable NoInClass
                y += 1              'Increments the variable on each cycle of the iteration
            Next
            If Con.State = ConnectionState.Closed Then      'Checks if the connections is closed
                Con.Open()                      'Sets the connection to open
            Else
                Con.Close()                      'Sets the connection to closed
                Con.Open()                      'Sets the connection to open
            End If
            Command = Con.CreateCommand         'Creates the settings for a new SQL Data Reader
            Command.CommandText = Cmd      'Sets the command text
            reader = Command.ExecuteReader   'Executes the command query and stores the results in the SQL data reader
            Do While reader.Read            'Checks if there is more data to be read
                If IsDBNull(reader.GetValue(0)) Then            'Checks if the reader contains a database null value
                    User(i).ResultP = 0                 'assigns the value 0 to variable if the reader contains a databse null value
                Else
                    User(i).ResultP = reader.GetValue(0)        'Assigns the value inside the reader to the variable
                End If
                If IsDBNull(reader.GetValue(1)) Then            'Checks if the reader contains a database null value
                    User(i).ResultS = 0                 'assigns the value 0 to variable if the reader contains a databse null value
                Else
                    User(i).ResultS = reader.GetValue(1)        'Assigns the value inside the reader to the variable
                End If
                If IsDBNull(reader.GetValue(2)) Then            'Checks if the reader contains a database null value
                    User(i).ResultW = 0                 'assigns the value 0 to variable if the reader contains a databse null value
                Else
                    User(i).ResultW = reader.GetValue(2)        'Assigns the value inside the reader to the variable
                End If
                If IsDBNull(reader.GetValue(3)) Then            'Checks if the reader contains a database null value
                    User(i).ResultI = 0                 'assigns the value 0 to variable if the reader contains a databse null value
                Else
                    User(i).ResultI = reader.GetValue(3)        'Assigns the value inside the reader to the variable
                End If
                User(i).Invigilator = reader.GetValue(4)        'Assigns the value inside the reader to the variable
                User(i).ClassName = reader.GetValue(5)        'Assigns the value inside the reader to the variable                                         'Assigns the value of i to n
                i += 1                                          'Increments i by 1
            Loop
            n = 0       'Assigns the value of 0 to n
            If Con.State = ConnectionState.Closed Then      'Checks if the connections is closed
                Con.Open()                      'Sets the connection to open
            Else
                Con.Close()                      'Sets the connection to closed
                Con.Open()                      'Sets the connection to open
            End If
            Command = Con.CreateCommand
            Command.CommandText = "SELECT COUNT(*) FROM TEST"
            TotalRows = CInt(Command.ExecuteScalar)
            ReDim user1(TotalRows)
            Command = Con.CreateCommand         'Creates the settings for a new SQL Data Reader
            Command.CommandText = Cmd1      'Sets the command text
            reader = Command.ExecuteReader   'Executes the command query and stores the results in the SQL data reader
            Do While reader.Read            'Checks if there is more data to be read
                If IsDBNull(reader.GetValue(0)) Then            'Checks if the reader contains a database null value
                    user1(n).result = "N/A"                 'assigns the value Failed to variable if the reader contains a databse null value
                Else
                    user1(n).result = reader.GetValue(0)        'Assigns the value inside the reader to the variable
                End If
                user1(n).ClassName = reader.GetValue(1)
                n += 1          'Increments n by 1
                If n = TotalRows Then       'Checks if n is equal to the totalrows so that an error doesnt occur of too many values being read
                    Exit Do
                End If
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)         'Sends an error message to the user if an error occurs
        Finally
            Con.Close()                 'Sets the connection to close
        End Try
        CalculateResults(ClassRequired, NoInClass, NoOfRowsClass, User, TotalRows, user1)
    End Sub

    Sub CalculateResults(ByVal ClassRequired() As String, ByVal NoInClass() As Integer, ByVal NoOfRowsClass As Integer, ByRef user() As Data, ByVal totalrows As Integer, ByRef user1() As Data1)
        Dim y, n As Integer
        Try
            y = (user.Length - 1)
            For i = 0 To (user1.Length - 1)
                For Each m In ClassRequired
                    If user1(i).ClassName = m Then
                        n = ClassRequired.IndexOf(ClassRequired, m)
                        If user1(i).result = "Passed" Then  'Checks if the variable is equal to passed
                            user(n).passFailratio += 1      'If the variable is equal it increments PassFailRatio by 1
                        ElseIf user1(i).result = "N/A" Then
                            user(n).NotTaken += 1
                        End If
                    End If
                Next
            Next
            For i = 0 To n
                If ((NoInClass(i) * 4) - user(i).NotTaken) <> 0 Then
                    user(i).passFailratio /= ((NoInClass(i) * 4) - user(i).NotTaken)         'Sets the variable to itslef dividedby the totalrows
                    user(i).passFailratio *= 100                'Multiplies the variable by 100 so that it is a percentage
                    user(i).passFailratio = Math.Round(user(i).passFailratio, 1)    'Rounds the variable to 1 decimal place
                End If
                If Not user(i).passFailratio = Nothing Then
                    LbxPassFailRatio.Items.Add(user(i).passFailratio & "%")         'If the PassFailRatio doesnt equal nothing it adds the variable to the Listbox
                Else
                    LbxPassFailRatio.Items.Add("N/A")
                End If
            Next
            For i = 0 To (user.Length - 1)
                For Each m In ClassRequired
                    If user(i).ClassName = m Then
                        n = ClassRequired.IndexOf(ClassRequired, m)
                        user(n).AverageResult += user(i).ResultP + user(i).ResultS + user(i).ResultW + user(i).ResultI       'Adds the result of each student in the class together
                    End If
                Next
            Next
            For i = 0 To n
                user(i).AverageResult /= (NoInClass(i) * 4)          'Divides the variable by the variable by the number of students in the class multipulied by the number of tests
                user(i).AverageResult = Math.Round(user(i).AverageResult, 1)        'Rounds the variable to 1 decimal place
                If Not IsNothing(user(i).AverageResult) Then            'Checks that the variable doesnt equal nothing
                    LbxAveargeResult.Items.Add(user(i).AverageResult & "%")     'Adds the value of the variable to the listbox
                End If
            Next
            For m = 0 To (user.Length - 1)
                If Not (user(m).ClassName = Nothing And user(m).Invigilator = Nothing) Then     'Checks that the classname and invigilator variables dont equal nothing
                    If user(m).ClassName <> user(y).ClassName Then
                        LbxClass.Items.Add(user(m).ClassName)       'Adds the classname variable value to the listbox
                        LbxInvigalator.Items.Add(user(m).Invigilator)       'Adds the invigilator variable value to the listbox
                        y = m
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)      'Sends an error message to the user
        End Try
    End Sub
End Class