Imports System.Data.SqlClient
Module freqFuncts
    Public Function sqlCheckOneCol(ByVal con As SqlConnection, ByVal cmd As SqlCommand, ByVal storagePlace As List(Of String))
        cmd.CommandType = CommandType.Text
        cmd.Connection = con
        con.Open()
        'adding values from the 1 column into the list
        Using sdr As SqlDataReader = cmd.ExecuteReader()
            If sdr.HasRows Then
                While sdr.Read
                    storagePlace.Add(sdr.GetString(0))
                End While
            End If
            sdr.Read()
        End Using
        con.Close()
    End Function
    Public Function sqlCheckTwoCols(ByVal con As SqlConnection, ByVal cmd As SqlCommand, ByVal storagePlace As List(Of String))
        cmd.CommandType = CommandType.Text
        cmd.Connection = con
        con.Open()
        'adding values from the 2 columns into the list
        Using sdr As SqlDataReader = cmd.ExecuteReader()
            If sdr.HasRows Then
                While sdr.Read
                    For index As Integer = 0 To 1
                        storagePlace.Add(sdr.GetString(index))
                    Next
                End While
            End If
            sdr.Read()
        End Using
        con.Close()
    End Function
    Public Function outputError(ByVal seg As String, ByVal id As Integer, ByVal compIdx As Integer, ByVal queueIdx As Integer, ByVal errMsg As String)
        If seg = "MSH" Then
            Dim outputRow As String() = New String() {id, "MSH", "Message Header", MSH.mshComponents(compIdx), MSH.msh(queueIdx), "", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), errMsg}
            parserInterface.DataGridView1.Rows.Add(outputRow)
            MSH.toBeParsedMsh.Enqueue(outputRow)
        End If
        If seg = "MSA" Then
            Dim outputRow As String() = New String() {id, "MSA", "Message Acknowledgement", MSA.msaComponents(compIdx), MSA.msa(queueIdx), "", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), errMsg}
            parserInterface.DataGridView1.Rows.Add(outputRow)
            MSA.toBeParsedMsa.Enqueue(outputRow)
        End If
        If seg = "NK1" Then
            Dim outputRow As String() = New String() {id, "NK1", "Next of Kin", NK1.nk1Components(compIdx), NK1.nk1(queueIdx), "", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), errMsg}
            parserInterface.DataGridView1.Rows.Add(outputRow)
            NK1.toBeParsedNk1.Enqueue(outputRow)
        End If
        If seg = "PV1" Then
            Dim outputRow As String() = New String() {id, "PV1", "Patient Visit", PV1.pv1Components(compIdx), PV1.pv1(queueIdx), "", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), errMsg}
            parserInterface.DataGridView1.Rows.Add(outputRow)
            PV1.toBeParsedPv1.Enqueue(outputRow)
        End If
    End Function
    Public Function outputConverted(ByVal seg As String, ByVal id As Integer, ByVal compIdx As Integer, ByVal queueIdx As Integer, ByVal convertedValue As String)
        If seg = "MSH" Then
            Dim outputRow As String() = New String() {id, "MSH", "Message Header", MSH.mshComponents(compIdx), MSH.msh(queueIdx), convertedValue, DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ""}
            parserInterface.DataGridView1.Rows.Add(outputRow)
            MSH.toBeParsedMsh.Enqueue(outputRow)
        End If
        If seg = "MSA" Then
            Dim outputRow As String() = New String() {id, "MSA", "Message Acknowledgement", MSA.msaComponents(compIdx), MSA.msa(queueIdx), convertedValue, DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ""}
            parserInterface.DataGridView1.Rows.Add(outputRow)
            MSA.toBeParsedMsa.Enqueue(outputRow)
        End If
        If seg = "NK1" Then
            Dim outputRow As String() = New String() {id, "NK1", "Next of Kin", NK1.nk1Components(compIdx), NK1.nk1(queueIdx), convertedValue, DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ""}
            parserInterface.DataGridView1.Rows.Add(outputRow)
            NK1.toBeParsedNk1.Enqueue(outputRow)
        End If
        If seg = "PV1" Then
            Dim outputRow As String() = New String() {id, "PV1", "Patient Visit", PV1.pv1Components(compIdx), PV1.pv1(queueIdx), convertedValue, DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ""}
            parserInterface.DataGridView1.Rows.Add(outputRow)
            PV1.toBeParsedPv1.Enqueue(outputRow)
        End If
    End Function
    Public Function comboBoxIf(ByVal input As Queue(Of String()))
        For Each info In input
            parserInterface.DataGridView1.Rows.Add(info)
        Next
    End Function
    Public Function comboBoxElse(ByVal input As Queue(Of String()))
        For Each info In input
            If info(4) = "" Then
                Continue For
            Else
                parserInterface.DataGridView1.Rows.Add(info)
            End If
        Next
    End Function
End Module
