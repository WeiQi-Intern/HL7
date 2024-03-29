﻿Imports System.Data.SqlClient
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

        '0 set ID - OBX - O;SI
        dataTypeSIoptional(id, obx(1), "OBX", 0, "Invalid set ID")

        '1 value type - O;ID
        Dim vt As New List(Of String)
        Using con As SqlConnection = New SqlConnection(constr)
            Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'valueType'")
                sqlCheckTwoCols(con, cmd, vt)
            End Using
        End Using
        If vt.Contains(obx(2)) Then
            outputConverted("OBX", id, 1, 2, vt(vt.IndexOf(obx(2)) + 1))
        ElseIf obx(2) = "" Then
            outputConverted("OBX", id, 1, 2, "")
        Else
            outputError("OBX", id, 1, 2, "Invalid value type")
        End If

        '2 Observation identifier - R;CWE
        usageRequired(id, obx(3), "OBX", 2, "Observation identifier is REQUIRED")

        '3 Observation sub ID ~ 9 Nature of abnormal test - O
        For idx As Integer = 3 To 9
            outputConverted("OBX", id, idx, idx + 1, "")
        Next

        '10 Observation result status - R;ID
        If obx(11) = "" Then
            outputError("OBX", id, 10, 11, "Observation result status is REQUIRED")
        Else
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
        End If

        '11 Effective date of reference range ~ 12 user defined access checks - O
        For idx As Integer = 10 To 11
            outputConverted("OBX", id, idx, idx + 1, "")
        Next

        '13 Date/time of the observation - O;DTM
        checkDTMoptional(id, obx(14), "OBX", 13)

        '14 Date/time of the obervation ~ Date/time of the analysis - O
        For idx As Integer = 14 To 18
            outputConverted("OBX", id, idx, idx + 1, "")
        Next
    End Function
End Class
