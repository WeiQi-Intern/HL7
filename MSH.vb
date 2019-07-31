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

        'link to sql
        Dim constr As String = "Data Source=LENOVO-330-VN6F\SQLEXPRESS;Initial Catalog=importFromExcel;user id=sa;password=111"

        '1. check usage: Required/Optional, 2. check data type, 3. then check values
        'add one to comment index to get in queue

        '0 Encoding characters - R;ST
        If msh(1) <> "^~\&" Then
            outputError("MSH", id, 0, 1, "Invalid delimiters.Should be ^~|&")
        Else
            outputConverted("MSH", id, 0, 1, "")
        End If

        '1 Sending application ~ 4 Receiving facility - O;HD
        For idx As Integer = 1 To 4
            Dim tempValue As String = ""
            tempValue = msh(idx + 1)
            Dim tempArray As String() = tempValue.Split(New Char() {"^"}, StringSplitOptions.RemoveEmptyEntries)
            Dim idType As New List(Of String)
            If tempArray.Length = 3 Then
                Using con As SqlConnection = New SqlConnection(constr)
                    Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM generalDTonly WHERE DataTypePart = 'HD3'")
                        sqlCheckTwoCols(con, cmd, idType)
                    End Using
                End Using
                If idType.Contains(tempArray(2)) Then
                    outputConverted("MSH", id, idx, idx + 1, tempArray(0) + ", " + tempArray(1) + ", " + idType(idType.IndexOf(msh(idx + 1)) + 2))
                Else
                    outputError("MSH", id, idx, idx + 1, "Invalid universal ID type")
                End If
            Else
                If tempArray.Length = 0 Or tempArray.Length = 1 Or tempArray.Length = 2 Then
                    outputConverted("MSH", id, idx, idx + 1, "")
                Else
                    outputError("MSH", id, idx, idx + 1, "Invalid data structure, maximum 3 sub-components.")
                End If
            End If
        Next

        ' 5 Date/time of message - R;DTM
        checkDTMrequired(id, msh(6), "MSH", 5)

        '6 Security - O;ST
        outputConverted("MSH", id, 6, 7, "")

        '7 Message Type - R;MSG
        If msh(8) <> "" Then
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
        Else
            outputError("MSH", id, 7, 8, "Message type is REQUIRED")
        End If

        '8 Message Control ID - R;ST
        usageRequired(id, msh(9), "MSH", 8, "Message control ID is REQUIRED")

        '9 Processing ID - R;PT
        If msh(10) <> "" Then
            Dim tempID As String = msh(10)
            Dim idArray As String() = tempID.Split(New Char() {"^"}, StringSplitOptions.RemoveEmptyEntries)
            If idArray.Length = 1 Then
                Dim processing1 As New List(Of String)
                Using con As SqlConnection = New SqlConnection(constr)
                    Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM generalDTonly WHERE DataTypePart = 'PT1'")
                        sqlCheckTwoCols(con, cmd, processing1)
                    End Using
                End Using
                If processing1.Contains(msh(10)) Then
                    outputConverted("MSH", id, 9, 10, processing1(processing1.IndexOf(msh(10)) + 1))
                Else
                    outputError("MSH", id, 9, 10, "Invalid processing ID")
                End If
            ElseIf idArray.Length = 2 Then
                Dim processing2 As New List(Of String)
                Using con As SqlConnection = New SqlConnection(constr)
                    Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM generalDTonly WHERE DataTypePart = 'PT1'")
                        sqlCheckTwoCols(con, cmd, processing2)
                    End Using
                End Using
                If processing2.Contains(idArray(0)) Then
                    Dim processing3 As New List(Of String)
                    Using con As SqlConnection = New SqlConnection(constr)
                        Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM generalDTonly WHERE DataTypePart = 'PT2'")
                            sqlCheckTwoCols(con, cmd, processing3)
                        End Using
                    End Using
                    If processing3.Contains(idArray(1)) Then
                        outputConverted("MSH", id, 9, 10, processing2(processing2.IndexOf(idArray(0)) + 1) + ", " + processing3(processing3.IndexOf(idArray(1)) + 1))
                    Else
                        outputError("MSH", id, 9, 10, "Invalid processing ID")
                    End If
                End If
            Else
                outputError("MSH", id, 9, 10, "Invalid processing ID structure")
            End If
        Else
            outputError("MSH", id, 9, 10, "Processing ID is REQUIRED")
        End If

        '10 Version ID - R;VID
        usageRequired(id, msh(11), "MSH", 10, "Version ID is REQUIRED")

        '11 Sequence number ~ 12 Continuation Pointer [excluded]
        For idx As Integer = 11 To 12
            outputConverted("MSH", id, idx, idx + 1, "")
        Next

        '13 Accept acknowledgement type ~ 14 Application acknowledgement type - O;ID
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

        '15 Country Code - O;ID
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

        '16 Character Set ~ 21 Receiving responsible organization -O
        For idx As Integer = 16 To 21
            outputConverted("MSH", id, idx, idx + 1, "")
        Next
    End Function
End Class
