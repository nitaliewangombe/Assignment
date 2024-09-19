Imports System.Data.OleDb

Module Module1
    Sub Main()
        Dim dbPath As String = "C:\Users\hp\OneDrive\Documents\AssignmentDb.accdb"

        ' Create a new Access database
        CreateDatabase(dbPath)
        CreateTables(dbPath)

        Console.WriteLine("Database and tables created successfully.")
    End Sub

    Sub CreateDatabase(ByVal dbPath As String)
        ' Create the Access database
        Dim accessEngine As Object = CreateObject("Access.Application")
        accessEngine.NewCurrentDatabase(dbPath)
        accessEngine.Quit()
    End Sub

    Sub CreateTables(ByVal dbPath As String)
        Using connection As New OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};")
            connection.Open()

            ' Create TblEmployee
            Dim createEmployeeTable As String = "
                CREATE TABLE TblEmployee (
                    EmployeeNumber SHORTTEXT,
                    EmployeeName SHORTTEXT,
                    Age INTEGER,
                    Gender SHORTTEXT,
                    Salary SHORTTEXT,
                    EmployeeDate DATETIME,
                    MaritalStatus SHORTTEXT,
                    EmployeeStatus SHORTTEXT
                )"
            Using command As New OleDbCommand(createEmployeeTable, connection)
                command.ExecuteNonQuery()
            End Using

            ' Create TblDepartment
            Dim createDepartmentTable As String = "
                CREATE TABLE TblDepartment (
                    DepartmentNumber TEXT,
                    DepartmentName TEXT,
                    PersonInCharge TEXT
                )"
            Using command As New OleDbCommand(createDepartmentTable, connection)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub
End Module