Imports System.IO
Imports System.Data.Sql
Imports System.Data.SqlClient
Public Class PopUP

    Structure ClassTable
        Dim ClassName, Invigilator, Trainer, AdminId As String
    End Structure

    Structure Student
        Dim BSCID, FirstName, Surname, Username, Password, RegDate, CompDate, Qualification, CertificateDate, ClassName, ResultP, ResultS, ResultW, ResultI As String
    End Structure

    Structure Test
        Dim BSCID, UnitCode, Testpaper, Result, Attempts, ResultofAttempt1, ResultofAttempt2, ResultofAttempt3, TestDate, StartDate, TimeSinceLastAttempt As String
    End Structure

    Structure Unit
        Dim UnitCode, Description, Syllabus As String
    End Structure

    Structure DDL
        Dim Student As String
        Dim ClassTable As String
        Dim Unit As String
        Dim Test As String
    End Structure

    Private Sub btnYes_Click(sender As Object, e As EventArgs) Handles btnYes.Click
        GetData() 'Calls the GetData Procedure
        MessageBox.Show("Data Has been Archived")
    End Sub

    Sub GetData()
        Dim TestTable() As Test
        Dim StudentTable() As Student
        Dim ClassTableData() As ClassTable
        Dim UnitTable() As Unit
        Dim TotalRows As Integer
        Dim i As Integer
        Dim Con As New SqlConnection(ConString)
        Dim Cmd As String = "SELECT * FROM Test"
        Dim Cmd1 As String = "SELECT * FROM Student"
        Dim Cmd2 As String = "SELECT * FROM Class"
        Dim cmd3 As String = "SELECT * FROM Unit"
        Dim Command As New SqlCommand
        Dim reader As SqlDataReader
        i = 0
        Try
            If Con.State = ConnectionState.Closed Then
                Con.Open()
            End If
            Command = Con.CreateCommand
            Command.CommandText = "SELECT COUNT(*) FROM Test"
            TotalRows = CInt(Command.ExecuteScalar)
            ReDim TestTable(TotalRows - 1)
            Command = Con.CreateCommand
            Command.CommandText = Cmd
            reader = Command.ExecuteReader
            Do While reader.Read
                TestTable(i).BSCID = reader.GetValue(0)
                TestTable(i).UnitCode = reader.GetValue(1)
                TestTable(i).Testpaper = reader.GetValue(2)
                If IsDBNull(reader.GetValue(3)) Then
                    TestTable(i).Result = "Failed"
                Else
                    TestTable(i).Result = CStr(reader.GetValue(3))
                End If
                If IsDBNull(reader.GetValue(4)) Then
                    TestTable(i).Attempts = "0"
                Else
                    TestTable(i).Attempts = CStr(reader.GetValue(4))
                End If
                If IsDBNull(reader.GetValue(5)) Then
                    TestTable(i).ResultofAttempt1 = "0"
                Else
                    TestTable(i).ResultofAttempt1 = CStr(reader.GetValue(5))
                End If
                If IsDBNull(reader.GetValue(6)) Then
                    TestTable(i).ResultofAttempt2 = "0"
                Else
                    TestTable(i).ResultofAttempt2 = CStr(reader.GetValue(6))
                End If
                If IsDBNull(reader.GetValue(7)) Then
                    TestTable(i).ResultofAttempt3 = "0"
                Else
                    TestTable(i).ResultofAttempt3 = CStr(reader.GetValue(7))
                End If
                TestTable(i).TestDate = reader.GetValue(8)
                TestTable(i).StartDate = reader.GetValue(9)
                If IsDBNull(reader.GetValue(10)) Then
                    TestTable(i).TimeSinceLastAttempt = CStr(Date.Now)
                Else
                    TestTable(i).TimeSinceLastAttempt = reader.GetValue(10)
                End If
                i += 1
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Con.Close()
        End Try
        i = 0
        Try
            If Con.State = ConnectionState.Closed Then
                Con.Open()
            End If
            Command = Con.CreateCommand
            Command.CommandText = "SELECT COUNT(*) FROM Student"
            TotalRows = CInt(Command.ExecuteScalar)
            ReDim StudentTable(TotalRows - 1)
            Command = Con.CreateCommand
            Command.CommandText = Cmd1
            reader = Command.ExecuteReader
            Do While reader.Read
                StudentTable(i).BSCID = reader.GetValue(0)
                StudentTable(i).FirstName = reader.GetValue(1)
                StudentTable(i).Surname = reader.GetValue(2)
                StudentTable(i).Username = reader.GetValue(3)
                StudentTable(i).Password = reader.GetValue(4)
                StudentTable(i).RegDate = CStr(reader.GetValue(5))
                StudentTable(i).CompDate = CStr(reader.GetValue(6))
                If IsDBNull(reader.GetValue(7)) Then
                    StudentTable(i).Qualification = "Fail"
                Else
                    StudentTable(i).Qualification = reader.GetValue(7)
                End If
                StudentTable(i).CertificateDate = CStr(reader.GetValue(8))
                StudentTable(i).ClassName = reader.GetValue(9)
                If IsDBNull(reader.GetValue(10)) Then
                    StudentTable(i).ResultP = "0"
                Else
                    StudentTable(i).ResultP = CStr(reader.GetValue(10))
                End If
                If IsDBNull(reader.GetValue(11)) Then
                    StudentTable(i).ResultS = "0"
                Else
                    StudentTable(i).ResultS = CStr(reader.GetValue(11))
                End If
                If IsDBNull(reader.GetValue(12)) Then
                    StudentTable(i).ResultW = "0"
                Else
                    StudentTable(i).ResultW = CStr(reader.GetValue(12))
                End If
                If IsDBNull(reader.GetValue(13)) Then
                    StudentTable(i).ResultI = "0"
                Else
                    StudentTable(i).ResultI = CStr(reader.GetValue(13))
                End If
                i += 1
            Loop
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            Con.Close()
        End Try
        i = 0
        Try
            If Con.State = ConnectionState.Closed Then
                Con.Open()
            End If
            Command = Con.CreateCommand
            Command.CommandText = "SELECT COUNT(*) FROM Class"
            TotalRows = CInt(Command.ExecuteScalar)
            ReDim ClassTableData(TotalRows - 1)
            Command = Con.CreateCommand
            Command.CommandText = Cmd2
            reader = Command.ExecuteReader
            Do While reader.Read
                ClassTableData(i).ClassName = reader.GetValue(0)
                ClassTableData(i).Invigilator = reader.GetValue(1)
                ClassTableData(i).Trainer = reader.GetValue(2)
                If IsDBNull(reader.GetValue(3)) Then
                    ClassTableData(i).AdminId = 0
                Else
                    ClassTableData(i).AdminId = reader.GetValue(3)
                End If
                i += 1
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Con.Close()
        End Try
        i = 0
        Try
            If Con.State = ConnectionState.Closed Then
                Con.Open()
            End If
            Command = Con.CreateCommand
            Command.CommandText = "SELECT COUNT(*) FROM Unit"
            TotalRows = CInt(Command.ExecuteScalar)
            ReDim UnitTable(TotalRows - 1)
            Command = Con.CreateCommand
            Command.CommandText = cmd3
            reader = Command.ExecuteReader
            Do While reader.Read
                UnitTable(i).UnitCode = reader.GetValue(0)
                UnitTable(i).Description = reader.GetValue(1)
                UnitTable(i).Syllabus = reader.GetValue(2)
                i += 1
            Loop
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Con.Close()
        End Try
        ArchiveData(ClassTableData, StudentTable, TestTable, UnitTable)
        ClearDatabase()
        Me.Close()
    End Sub

    Sub ArchiveData(ByVal ClassTableData() As ClassTable, ByVal StudentTable() As Student, ByVal TestTable() As Test, ByVal UnitTable() As Unit)
        Try
            Dim temp As String
            Dim FileName As String
            Dim Writer As BinaryWriter
            temp = InputBox("Please Enter a Year for the Acidemic Year")
            FileName = "Class" & temp & ".dat"
            Writer = New BinaryWriter(File.Open(FileName, FileMode.Create))
            For i = 0 To (ClassTableData.Length - 1)
                Writer.Write(ClassTableData(i).ClassName)
                Writer.Write(ClassTableData(i).Invigilator)
                Writer.Write(ClassTableData(i).Trainer)
                Writer.Write(ClassTableData(i).AdminId)
            Next
            FileName = "Student" & temp & ".dat"
            Writer = New BinaryWriter(File.Open(FileName, FileMode.Create))
            For i = 0 To (StudentTable.Length - 1)
                Writer.Write(StudentTable(i).BSCID)
                Writer.Write(StudentTable(i).FirstName)
                Writer.Write(StudentTable(i).Surname)
                Writer.Write(StudentTable(i).Username)
                Writer.Write(StudentTable(i).RegDate)
                Writer.Write(StudentTable(i).CompDate)
                Writer.Write(StudentTable(i).Qualification)
                Writer.Write(StudentTable(i).CertificateDate)
                Writer.Write(StudentTable(i).ClassName)
                Writer.Write(StudentTable(i).ResultP)
                Writer.Write(StudentTable(i).ResultS)
                Writer.Write(StudentTable(i).ResultW)
                Writer.Write(StudentTable(i).ResultI)
            Next
            FileName = "Test" & temp & ".dat"
            Writer = New BinaryWriter(File.Open(FileName, FileMode.Create))
            For i = 0 To (TestTable.Length - 1)
                Writer.Write(TestTable(i).BSCID)
                Writer.Write(TestTable(i).UnitCode)
                Writer.Write(TestTable(i).Testpaper)
                Writer.Write(TestTable(i).Result)
                Writer.Write(TestTable(i).Attempts)
                Writer.Write(TestTable(i).ResultofAttempt1)
                Writer.Write(TestTable(i).ResultofAttempt2)
                Writer.Write(TestTable(i).ResultofAttempt3)
                Writer.Write(TestTable(i).TestDate)
                Writer.Write(TestTable(i).StartDate)
                Writer.Write(TestTable(i).TimeSinceLastAttempt)
            Next
            FileName = "Unit" & temp & ".dat"
            Writer = New BinaryWriter(File.Open(FileName, FileMode.Create))
            For i = 0 To (UnitTable.Length - 1)
                Writer.Write(UnitTable(i).UnitCode)
                Writer.Write(UnitTable(i).Description)
                Writer.Write(UnitTable(i).Syllabus)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Sub ClearDatabase()
        Dim Con As New SqlConnection(ConString)
        Dim Command As New SqlCommand
        Dim Script As DDL
        Dim Filename As String = "DDLScript.Dat"
        Dim reader As BinaryReader = New BinaryReader(File.Open(Filename, FileMode.Open))
        Try
            Script.ClassTable = reader.ReadString
            Script.Student = reader.ReadString
            Script.Test = reader.ReadString
            Script.Unit = reader.ReadString
            If Con.State = ConnectionState.Closed Then
                Con.Open()
            End If
            Command = Con.CreateCommand
            Command.CommandText = "DROP TABLE Test"
            Command.ExecuteNonQuery()
            Command = Con.CreateCommand
            Command.CommandText = "DROP TABLE Unit"
            Command.ExecuteNonQuery()
            Command = Con.CreateCommand
            Command.CommandText = "DROP TABLE Student"
            Command.ExecuteNonQuery()
            Command = Con.CreateCommand
            Command.CommandText = "DROP TABLE Class"
            Command.ExecuteNonQuery()
            Command = Con.CreateCommand
            Command.CommandText = Script.ClassTable
            Command.ExecuteNonQuery()
            Command = Con.CreateCommand
            Command.CommandText = Script.Student
            Command.ExecuteNonQuery()
            Command = Con.CreateCommand
            Command = Con.CreateCommand
            Command.CommandText = Script.Unit
            Command.ExecuteNonQuery()
            Command.CommandText = Script.Test
            Command.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub btnNo_Click(sender As Object, e As EventArgs) Handles btnNo.Click
        Me.Close()
    End Sub
End Class