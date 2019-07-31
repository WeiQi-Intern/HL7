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

        '0 Set ID - PID - O;SI
        dataTypeSIoptional(id, pid(1), "PID", 0, "Invalid set ID")

        '1 Patient ID - O;CX
        outputConverted("PID", id, 1, 2, "")

        '2 Patient identifier list - R;CX
        usageRequired(id, pid(3), "PID", 2, "Patient identifier list is REQUIRED")

        '3 Alternate Patient ID – PID - O;CX
        outputConverted("PID", id, 3, 4, "")

        '4 Patient Name - R;XPN
        usageRequired(id, pid(5), "PID", 4, "Patient name is REQUIRED")

        '5 Mother's Maiden Name ~ 6 Date/Time of Birth - O [excluded]
        For idx As Integer = 5 To 6
            outputConverted("PID", id, idx, idx + 1, "")
        Next

        '7 Gender - O;IS 
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

        '8 Patient Alias - O;XPN
        outputConverted("PID", id, 8, 9, "")

        '9 Race - O;CWE
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

        '10 Patient address ~ 14 Primary language - O
        For idx As Integer = 10 To 14
            outputConverted("PID", id, idx, idx + 1, "")
        Next

        '15 Marital status - O;CWE
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

        '16 Religion - O;CWE
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

        '17 Patient account number ~ 20 mother's identifier - O
        For idx As Integer = 17 To 20
            outputConverted("PID", id, idx, idx + 1, "")
        Next

        '21 Ethnic group - O;CWE
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

        '22 Birth place - O;XAD
        outputConverted("PID", id, 22, 23, "")

        '23 Multiple birth indicator - O;ID
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

        '24 Birth order ~ 27 Nationality - O
        For idx As Integer = 24 To 27
            outputConverted("PID", id, idx, idx + 1, "")
        Next

        '28 Patient death date and time
        checkDTMoptional(id, pid(29), "PID", 28)

        '29 Patient death indicator
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
