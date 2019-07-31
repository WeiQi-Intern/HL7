Imports System.Data.SqlClient
Public Class NK1
    Public Shared nk1 As Queue(Of String) = New Queue(Of String)
    Public Shared nk1Components As String() = {"Set ID", "Name", "Relationship", "Address", "Phone Number", "Business Phone Number", "Contact Role", "Start Date", "End Date", "Next of Kin/Associated Parties Job Title", "Next of Kin/Associated Parties Job Code/Class", "Next of Kin/Associated Parties Employee Number", "Organization Name - NK1", "Marital Status", "Administrative Sex", "Date/Time of Birth", "Living Dependency", "Ambulatory Status", "Citizenship", "Primary Language", "Living Arrangement", "Publicity Code", "Protection Indicator", "Student Indicator", "Religion", "Mother's Maiden Name", "Nationality", "Ethnic Group", "Contact Reason", "Contact Person's Name", "Contact Person's Telephone Number", "Contact Person's Address", "Next of Kin/Associated Party's Identifiers", "Job Status", "Race", "Handicap", "Contact Person Social Security Number"}
    Public Shared toBeParsedNk1 As Queue(Of String()) = New Queue(Of String())
    Public Shared Function parseNK1(ByVal message As String, ByVal id As Integer)
        'STEP 1: clear history
        nk1.Clear()
        toBeParsedNk1.Clear()
        parserInterface.ComboBox1.Refresh()

        'STEP 2: split info
        Dim rawNk1 As String() = message.Split(New Char() {"|"})
        'the last | will create an unwanted space, so remove it
        Dim indexOfLastElement As Integer = rawNk1.Count() - 1
        For i As Integer = 0 To (indexOfLastElement - 1)
            nk1.Enqueue(rawNk1(i))
        Next

        'STEP 3: check values

        'link to sql
        Dim constr As String = "Data Source=LENOVO-330-VN6F\SQLEXPRESS;Initial Catalog=importFromExcel;user id=sa;password=111"

        '0 Set ID - R;SI
        If nk1(1) = "" Then
            outputError("NK1", id, 0, 1, "Set ID is REQUIRED")
        Else
            dataTypeSIrequired(id, nk1(1), "NK1", 0, "Invalid set ID")
        End If

        '1 Name - O;XPN
        outputConverted("NK1", id, 1, 2, "")


        '2 relationship - O;CWE
        Dim rs As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'relationship'")
                sqlCheckTwoCols(con, cmd, rs)
            End Using
        End Using
        If rs.Contains(nk1(3)) Then
            outputConverted("NK1", id, 2, 3, rs(rs.IndexOf(nk1(3)) + 1))
        Else
            outputError("NK1", id, 2, 3, "Invalid relationship")
        End If

        '3 Address ~ 5 Business phone number - O
        For idx As Integer = 3 To 5
            outputConverted("NK1", id, idx, idx + 1, "")
        Next

        '6 Contact role - O;CWE
        Dim cr As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'contactRole'")
                sqlCheckTwoCols(con, cmd, cr)
            End Using
        End Using
        If cr.Contains(nk1(7)) Then
            outputConverted("NK1", id, 6, 7, cr(cr.IndexOf(nk1(7)) + 1))
        Else
            outputError("NK1", id, 6, 7, "Invalid contact role")
        End If

        '7 Start Date ~ 32 Next of Kin/Associated Party's Identifiers - O
        For idx As Integer = 7 To 32
            outputConverted("NK1", id, idx, idx + 1, "")
        Next

        '33 Job status - O;IS
        Dim jobSts As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'jobStatus'")
                sqlCheckTwoCols(con, cmd, jobSts)
            End Using
        End Using
        If jobSts.Contains(nk1(34)) Then
            outputConverted("NK1", id, 33, 34, jobSts(jobSts.IndexOf(nk1(34)) + 1))
        Else
            outputError("NK1", id, 33, 34, "Invalid contact role")
        End If

        '34 to 36 - O
        For idx As Integer = 34 To 36
            outputConverted("NK1", id, idx, idx + 1, "")
        Next
    End Function
End Class
