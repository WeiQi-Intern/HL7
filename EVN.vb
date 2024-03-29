﻿Imports System.Data.SqlClient
Public Class EVN
    Public Shared evn As Queue(Of String) = New Queue(Of String)
    Public Shared evnComponents As String() = {"Event Type Code", "Recorded Date/Time", "Date/Time Planned Event", "Event Reason Code", "Operator ID", "Event Occured", "Event Facility"}
    Public Shared toBeParsedEvn As Queue(Of String()) = New Queue(Of String())
    Public Shared Function parseEVN(ByVal message As String, ByVal id As Integer)
        'STEP 1        
        evn.Clear()
        toBeParsedEvn.Clear()
        parserInterface.ComboBox1.Refresh()

        'STEP 2
        Dim rawEvn As String() = message.Split(New Char() {"|"})
        'the last | will create an unwanted space, so remove it
        Dim indexOfLastElement As Integer = rawEvn.Count() - 1
        For i As Integer = 0 To (indexOfLastElement - 1)
            evn.Enqueue(rawEvn(i))
        Next

        'STEP 3: cross check values and convert if necessary
        'link to sql
        Dim constr As String = "Data Source=LENOVO-330-VN6F\SQLEXPRESS;Initial Catalog=importFromExcel;user id=sa;password=111"

        '0 Event type - O;ID
        Dim evntType As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE TableCode = '0003'")
                sqlCheckTwoCols(con, cmd, evntType)
            End Using
        End Using
        If evntType.Contains(evn(1)) Then
            outputConverted("EVN", id, 0, 1, evntType(evntType.IndexOf(evn(1)) + 1))
        Else
            outputError("EVN", id, 0, 1, "Invalid event type")
        End If

        '1 Recorded date/time - O;DTM
        checkDTMoptional(id, evn(2), "EVN", 1)

        '2 Date/time planned event [excluded]
        outputConverted("EVN", id, 2, 3, "")

        '3 Event reason code - O;IS
        Dim evntR As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE TableCode = '0062'")
                sqlCheckTwoCols(con, cmd, evntR)
            End Using
        End Using
        If evntR.Contains(evn(4)) Then
            outputConverted("EVN", id, 3, 4, evntR(evntR.IndexOf(evn(4)) + 1))
        Else
            outputError("EVN", id, 3, 4, "Invalid event reason code")
        End If

        '4 Operator ID - O;XCN ~ 6 Event Facility [excluded]
        For idx As Integer = 4 To 6
            outputConverted("EVN", id, idx, idx + 1, "")
        Next
    End Function
End Class
