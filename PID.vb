Imports System.Data.SqlClient
Public Class PID
    Public Shared pid As Queue(Of String) = New Queue(Of String)
    Public Shared pidComponents As String() = {"Set ID – PID", "Patient ID", "Patient Identifier List", "Alternate Patient ID – PID", "Patient Name", "Mother's Maiden Name", "Date/Time of Birth", "Sex", "Patient Alias", "Race", "Patient Address", "County Code", "Phone Number – Home", "Phone Number – Work", "Primary Language", "Marital Status", "Religion", "Patient Account Number", "SSN Number – Patient", "Driver's License Number – Patient", "Mother's Identifier", "Ethnic Group", "Birth Place", "Multiple Birth Indicator", "Birth Order", "Citizenship", "Veterans Military Status", "Nationality", "Patient Death Date and Time", "Patient Death Indicator"}
    Public Shared toBeParsedPid As Queue(Of String()) = New Queue(Of String())
    Public Shared Function parsePID(ByVal message As String, ByVal id As Integer)
        'STEP 1: clear history
        pid.Clear()
        toBeParsedPid.Clear()
        parserInterface.ComboBox1.Refresh()

        'STEP 2: split info
        Dim rawPid As String() = message.Split(New Char() {"|"})
        Dim indexOfLastElement As Integer = rawPid.Count() - 1
        For i As Integer = 0 To (indexOfLastElement - 1)
            pid.Enqueue(rawPid(i))
        Next

        'STEP 3: check with sql
        Dim constr As String = "Data Source=LENOVO-330-VN6F\SQLEXPRESS;Initial Catalog=importFromExcel;user id=sa;password=111"

        '0 to 6
        For idx As Integer = 0 To 6
            outputConverted("PID", id, idx, idx + 1, "")
        Next

        '8 gender
        Dim gender As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'administrativeSex'")
                sqlCheckTwoCols(con, cmd, gender)
            End Using
        End Using
        If gender.Contains(pid(8)) Then
            outputConverted("PID", id, 7, 8, gender(gender.IndexOf(pid(8)) + 1))
        Else
            outputError("PID", id, 7, 8, "Invalid gender")
        End If

        '9 
        outputConverted("PID", id, 8, 9, "")

        '10
        Dim race As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'race'")
                sqlCheckTwoCols(con, cmd, race)
            End Using
        End Using
        If race.Contains(pid(10)) Then
            outputConverted("PID", id, 9, 10, race(race.IndexOf(pid(10)) + 1))
        Else
            outputError("PID", id, 9, 10, "Invalid race")
        End If

        '10 to 14
        For idx As Integer = 10 To 14
            outputConverted("PID", id, idx, idx + 1, "")
        Next

        '16
        Dim maritalStatus As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'maritalStatus'")
                sqlCheckTwoCols(con, cmd, maritalStatus)
            End Using
        End Using
        If maritalStatus.Contains(pid(16)) Then
            outputConverted("PID", id, 15, 16, maritalStatus(maritalStatus.IndexOf(pid(16)) + 1))
        Else
            outputError("PID", id, 15, 16, "Invalid marital status")
        End If

        '17 religion
        Dim religion As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'religion'")
                sqlCheckTwoCols(con, cmd, religion)
            End Using
        End Using
        If religion.Contains(pid(17)) Then
            outputConverted("PID", id, 16, 17, religion(religion.IndexOf(pid(17)) + 1))
        Else
            outputError("PID", id, 16, 17, "Invalid religion")
        End If

        '17 to 20
        For idx As Integer = 17 To 20
            outputConverted("PID", id, idx, idx + 1, "")
        Next

        '22 ethnic group
        Dim eg As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'ethnicGroup'")
                sqlCheckTwoCols(con, cmd, eg)
            End Using
        End Using
        If eg.Contains(pid(22)) Then
            outputConverted("PID", id, 21, 22, eg(eg.IndexOf(pid(22)) + 1))
        Else
            outputError("PID", id, 21, 22, "Invalid ethnic group")
        End If

        '23
        outputConverted("PID", id, 22, 23, "")

        '24 birth indicator
        Dim bi As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'yesNoIndicator'")
                sqlCheckTwoCols(con, cmd, bi)
            End Using
        End Using
        If bi.Contains(pid(24)) Then
            outputConverted("PID", id, 23, 24, bi(bi.IndexOf(pid(24)) + 1))
        Else
            outputError("PID", id, 23, 24, "Invalid birth indicator")
        End If

        '25 to 29
        For idx As Integer = 24 To 28
            outputConverted("PID", id, idx, idx + 1, "")
        Next

        '30 patient death indicator
        Dim pdi As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'yesNoIndicator'")
                sqlCheckTwoCols(con, cmd, pdi)
            End Using
        End Using
        If bi.Contains(pid(30)) Then
            outputConverted("PID", id, 29, 30, pdi(pdi.IndexOf(pid(30)) + 1))
        Else
            outputError("PID", id, 29, 30, "Invalid birth indicator")
        End If
    End Function
End Class
