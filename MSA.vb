﻿Imports System.Data.SqlClient
Public Class MSA
    Public Shared msa As Queue(Of String) = New Queue(Of String)
    Public Shared msaComponents As String() = {"Acknowledgement Code", "Message Control ID", "Text Message", "Expected Sequence Number", "Delayed Acknowledgement Type", "Error Condition"}
    Public Shared toBeParsedMsa As Queue(Of String()) = New Queue(Of String())
    Public Shared Function parseMSA(ByVal message As String, ByVal id As Integer)
        'STEP 1: clear history
        msa.Clear()
        toBeParsedMsa.Clear()
        parserInterface.ComboBox1.Refresh()

        'STEP 2: split information (the last | will create an unwanted space, so remove it)
        Dim rawMsa As String() = message.Split(New Char() {"|"})
        Dim indexOfLastElement As Integer = rawMsa.Count() - 1
        For i As Integer = 0 To (indexOfLastElement - 1)
            msa.Enqueue(rawMsa(i))
        Next

        'link to sql
        Dim constr As String = "Data Source=LENOVO-330-VN6F\SQLEXPRESS;Initial Catalog=importFromExcel;user id=sa;password=111"

        '0 acknowledgment code - R;ID
        If msa(1) <> "" Then
            Dim ackCode As New List(Of String)
            Using con As SqlConnection = New SqlConnection(constr)
                Using cmd As SqlCommand = New SqlCommand("SELECT Code, Label FROM convertedValuesTable WHERE Type = 'acknowledgementCode'")
                    sqlCheckTwoCols(con, cmd, ackcode)
                End Using
            End Using
            If ackCode.Contains(msa(1)) Then
                outputConverted("MSA", id, 0, 1, ackCode(ackCode.IndexOf(msa(1)) + 1))
            Else
                outputError("MSA", id, 0, 1, "Invalid acknowledgement code")
            End If
        Else
            outputError("MSA", id, 0, 1, "Acknowledgement code is REQUIRED")
        End If

        '1 Message control ID - R;ST
        usageRequired(id, msa(2), "MSA", 1, "Message control ID is REQUIRED")

        '2 Text messsage - O;ST
        outputConverted("MSA", id, 2, 3, "")

        '3 Expected sequence number - O;NM
        dataTypeNM(id, msa(4), "MSA", 3, "Invalid expected seqeunce number")

        '4 Delayed acknowledgement type ~ 5 Error condition [excluded]
        For idx As Integer = 4 To 5
            outputConverted("MSA", id, idx, idx + 1, "")
        Next
    End Function
End Class
