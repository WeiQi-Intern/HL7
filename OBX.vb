Imports System.Data.SqlClient
Public Class OBX
    Public Shared obx As Queue(Of String) = New Queue(Of String)
    Public Shared obxComponents As String() = {"Set ID – OBX", "Value Type", "Observation Identifier", "Observation Sub-ID", "Observation Value", "Units", "References Range", "Abnormal Flags", "Probability", "Nature of Abnormal Test", "Observation Result Status", "Effective Date of Reference Range", "User Defined Access Checks", "Date/Time of the Observation", "Producer's ID", "Responsible Observer", "Observation Method", "Equipment Instance Identifier", "Date/Time of the Analysis"}
    Public Shared toBeParsedObx As Queue(Of String()) = New Queue(Of String())
    Public Shared Function parseOBX(ByVal message As String, ByVal id As Integer)
        'step1 
        obx.Clear()
        toBeParsedObx.Clear()
        parserInterface.ComboBox1.Refresh()

        'step2
        Dim rawObx As String() = message.Split(New Char() {"|"})
        Dim indexOfLastElement As Integer = rawObx.Count() - 1
        For i As Integer = 0 To (indexOfLastElement - 1)
            obx.Enqueue(rawObx(i))
        Next

        'step3
        'link to sql
        Dim constr As String = "Data Source=LENOVO-330-VN6F\SQLEXPRESS;Initial Catalog=importFromExcel;user id=sa;password=111"

        '0
        outputConverted("OBX", id, 0, 1, "")

        '1 value type
        Dim vt As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'valueType'")
                sqlCheckTwoCols(con, cmd, vt)
            End Using
        End Using
        If vt.Contains(obx(2)) Then
            outputConverted("OBX", id, 1, 2, vt(vt.IndexOf(obx(2)) + 1))
        Else
            outputError("OBX", id, 1, 2, "Invalid value type")
        End If

        '2 to 10
        For idx As Integer = 3 To 10
            outputConverted("OBX", id, idx - 1, idx, "")
        Next

        '11 observation result status
        Dim ors As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE TableCode = '0085'")
                sqlCheckTwoCols(con, cmd, ors)
            End Using
        End Using
        If ors.Contains(obx(11)) Then
            outputConverted("OBX", id, 10, 11, ors(ors.IndexOf(obx(11)) + 1))
        Else
            outputError("OBX", id, 10, 11, "Invalid observation result status")
        End If

        '12 to 19
        For idx As Integer = 12 To 19
            outputConverted("OBX", id, idx - 1, idx, "")
        Next
    End Function
End Class
