Imports System.Data.SqlClient
Public Class MSH
    'note that the index has a difference of 1 in msh vs mshComponents
    Public Shared msh As Queue(Of String) = New Queue(Of String)
    Public Shared mshComponents As String() = {"Encoding Characters", "Sending Application", "Sending Facility", "Receiving Application", "Receiving Facility", "Date/Time of Message", "Security", "Message Type", "Message Control ID", "Processing ID", "Version ID", "Sequence Number", "Continuation Pointer", "Accept Acknowledgement Type", "Application Acknowledgement Type", "Country Code", "Character Set", "Principal Language of Message", "Alternate Character Set Handling Scheme", "Message Profile Identifier", "Sending Responsible Organization", "Receiving Responsible Organization"}
    Public Shared toBeParsedMsh As Queue(Of String()) = New Queue(Of String())
    Public Shared Function parseMSH(ByVal message As String, ByVal id As Integer)
        'STEP 1: clear history
        msh.Clear()
        toBeParsedMsh.Clear()
        parserInterface.DataGridView1.Rows.Clear()
        parserInterface.ComboBox1.Refresh()

        'STEP 2: split information (the last | will create an unwanted space, so remove it)
        Dim rawMsh As String() = message.Split(New Char() {"|"})
        Dim indexOfLastElement As Integer = rawMsh.Count() - 1
        For i As Integer = 0 To (indexOfLastElement - 1)
            msh.Enqueue(rawMsh(i))
        Next

        'STEP 3: cross check values and convert if necessary

        '0 to 6
        For idx As Integer = 0 To 6
            outputConverted("MSH", id, idx, idx + 1, "")
        Next

        'link to sql
        Dim constr As String = "Data Source=LENOVO-330-VN6F\SQLEXPRESS;Initial Catalog=importFromExcel;user id=sa;password=111"

        '8 Message Type
        Dim checkStruc As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code FROM convertedValuesTable WHERE Type = 'messageStructure'")
                sqlCheckOneCol(con, cmd, checkStruc)
            End Using
        End Using
        'criteria 1: message structure must first be valid
        If checkStruc.Contains(msh(8)) Then
            Dim convertMsg As New List(Of String)
            Using con As SqlConnection = New SqlConnection(constr)
                Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'messageType'")
                    sqlCheckTwoCols(con, cmd, convertMsg)
                End Using
            End Using
            Dim convertEvnt As New List(Of String)
            Using con As SqlConnection = New SqlConnection(constr)
                Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'eventType'")
                    sqlCheckTwoCols(con, cmd, convertEvnt)
                    'ALERT: here got prob
                End Using
            End Using
            'criteria 2: message type must be valid
            If convertMsg.Contains(msh(8)(0) + msh(8)(1) + msh(8)(2)) Then
                Dim valuePos1 As Integer = convertMsg.IndexOf(msh(8)(0) + msh(8)(1) + msh(8)(2))
                Dim info1 As String = convertMsg(valuePos1 + 1)
                'criteria 3: event type must be valid
                If convertEvnt.Contains(msh(8)(4) + msh(8)(5) + msh(8)(6)) Then
                    Dim valuePos2 As Integer = convertEvnt.IndexOf(msh(8)(4) + msh(8)(5) + msh(8)(6))
                    Dim info2 As String = convertEvnt(valuePos2 + 1)
                    outputConverted("MSH", id, 7, 8, info1 + ", " + info2)
                Else
                    outputError("MSH", id, 7, 8, "Invalid event type")
                End If
            Else
                outputError("MSH", id, 7, 8, "Invalid message type")
            End If
        Else
            outputError("MSH", id, 7, 8, "Invalid message structure")
        End If
        outputConverted("MSH", id, 8, 9, "")

        '10 Processing ID
        Dim processingID As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'processingID'")
                sqlCheckTwoCols(con, cmd, processingID)
            End Using
        End Using
        If processingID.Contains(msh(10)) Then
            outputConverted("MSH", id, 9, 10, processingID(processingID.IndexOf(msh(10)) + 1))
        Else
            outputError("MSH", id, 9, 10, "Invalid processing ID")
        End If

        '11 to 13
        For idx As Integer = 11 To 13
            outputConverted("MSH", id, idx - 1, idx, "")
        Next

        '14 Accept acknowledgement type
        For idx As Integer = 14 To 15
            Dim ackType As New List(Of String)
            Using con As SqlConnection = New SqlConnection(constr)
                Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'acknowledgementType'")
                    sqlCheckTwoCols(con, cmd, ackType)
                End Using
            End Using
            If ackType.Contains(msh(idx)) Then
                outputConverted("MSH", id, idx - 1, idx, ackType(ackType.IndexOf(msh(idx)) + 1))
            Else
                If idx = 14 Then
                    outputError("MSH", id, 13, 14, "Invalid accept acknowldgement type")
                Else
                    outputError("MSH", id, 14, 15, "Invalid application acknowledgement type")
                End If
            End If
        Next

        '16 Country Code
        Dim countryCode As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'countryCode'")
                sqlCheckTwoCols(con, cmd, countryCode)
            End Using
        End Using
        If countryCode.Contains(msh(16)) Then
            outputConverted("MSH", id, 15, 16, countryCode(countryCode.IndexOf(msh(16)) + 1))
        Else
            outputError("MSH", id, 15, 16, "Invalid country code")
        End If

        '17 to 21
        For idx As Integer = 17 To 21
            outputConverted("MSH", id, idx, idx + 1, "")
        Next
    End Function
End Class
