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
            If parserInterface.showBlanksOnly = "yes" And MSH.msh(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSH.toBeParsedMsh.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSH.toBeParsedMsh.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSH.toBeParsedMsh.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (MSH.msh(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSH.toBeParsedMsh.Enqueue(outputRow)
            End If
        End If
        If seg = "MSA" Then
            Dim outputRow As String() = New String() {id, "MSA", "Message Acknowledgement", MSA.msaComponents(compIdx), MSA.msa(queueIdx), "", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), errMsg}
            If parserInterface.showBlanksOnly = "yes" And MSA.msa(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSA.toBeParsedMsa.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSA.toBeParsedMsa.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSA.toBeParsedMsa.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (MSA.msa(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSA.toBeParsedMsa.Enqueue(outputRow)
            End If
        End If
        If seg = "NK1" Then
            Dim outputRow As String() = New String() {id, "NK1", "Next of Kin", NK1.nk1Components(compIdx), NK1.nk1(queueIdx), "", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), errMsg}
            If parserInterface.showBlanksOnly = "yes" And NK1.nk1(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NK1.toBeParsedNk1.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NK1.toBeParsedNk1.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NK1.toBeParsedNk1.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (NK1.nk1(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NK1.toBeParsedNk1.Enqueue(outputRow)
            End If
        End If
        If seg = "PV1" Then
            Dim outputRow As String() = New String() {id, "PV1", "Patient Visit", PV1.pv1Components(compIdx), PV1.pv1(queueIdx), "", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), errMsg}
            If parserInterface.showBlanksOnly = "yes" And PV1.pv1(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PV1.toBeParsedPv1.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PV1.toBeParsedPv1.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PV1.toBeParsedPv1.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (PV1.pv1(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PV1.toBeParsedPv1.Enqueue(outputRow)
            End If
        End If
        If seg = "EVN" Then
            Dim outputRow As String() = New String() {id, "EVN", "Event Type", EVN.evnComponents(compIdx), EVN.evn(queueIdx), "", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), errMsg}
            If parserInterface.showBlanksOnly = "yes" And EVN.evn(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                EVN.toBeParsedEvn.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                EVN.toBeParsedEvn.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                EVN.toBeParsedEvn.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (EVN.evn(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                EVN.toBeParsedEvn.Enqueue(outputRow)
            End If
        End If
        If seg = "PID" Then
            Dim outputRow As String() = New String() {id, "PID", "Patient Identification", PID.pidComponents(compIdx), PID.pid(queueIdx), "", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), errMsg}
            If parserInterface.showBlanksOnly = "yes" And PID.pid(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PID.toBeParsedPid.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PID.toBeParsedPid.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PID.toBeParsedPid.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (PID.pid(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PID.toBeParsedPid.Enqueue(outputRow)
            End If
        End If
        If seg = "OBR" Then
            Dim outputRow As String() = New String() {id, "OBR", "Observation Reports", OBR.obrComponents(compIdx), OBR.obr(queueIdx), "", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), errMsg}
            If parserInterface.showBlanksOnly = "yes" And OBR.obr(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBR.toBeParsedObr.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBR.toBeParsedObr.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBR.toBeParsedObr.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (OBR.obr(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBR.toBeParsedObr.Enqueue(outputRow)
            End If
        End If
        If seg = "OBX" Then
            Dim outputRow As String() = New String() {id, "OBX", "Observation", OBX.obxComponents(compIdx), OBX.obx(queueIdx), "", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), errMsg}
            If parserInterface.showBlanksOnly = "yes" And OBX.obx(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBX.toBeParsedObx.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBX.toBeParsedObx.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBX.toBeParsedObx.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (OBX.obx(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBX.toBeParsedObx.Enqueue(outputRow)
            End If
        End If
        If seg = "NTE" Then
            Dim outputRow As String() = New String() {id, "NTE", "Notes and Comments", NTE.nteComponents(compIdx), NTE.nte(queueIdx), "", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), errMsg}
            If parserInterface.showBlanksOnly = "yes" And NTE.nte(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NTE.toBeParsedNte.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NTE.toBeParsedNte.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NTE.toBeParsedNte.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (NTE.nte(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NTE.toBeParsedNte.Enqueue(outputRow)
            End If
        End If
    End Function
    Public Function outputConverted(ByVal seg As String, ByVal id As Integer, ByVal compIdx As Integer, ByVal queueIdx As Integer, ByVal convertedValue As String)
        If seg = "MSH" Then
            Dim outputRow As String() = New String() {id, "MSH", "Message Header", MSH.mshComponents(compIdx), MSH.msh(queueIdx), convertedValue, DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ""}
            If parserInterface.showBlanksOnly = "yes" And MSH.msh(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSH.toBeParsedMsh.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSH.toBeParsedMsh.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSH.toBeParsedMsh.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (MSH.msh(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSH.toBeParsedMsh.Enqueue(outputRow)
            End If
        End If
        If seg = "MSA" Then
            Dim outputRow As String() = New String() {id, "MSA", "Message Acknowledgement", MSA.msaComponents(compIdx), MSA.msa(queueIdx), convertedValue, DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ""}
            If parserInterface.showBlanksOnly = "yes" And MSA.msa(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSA.toBeParsedMsa.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSA.toBeParsedMsa.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSA.toBeParsedMsa.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (MSA.msa(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                MSA.toBeParsedMsa.Enqueue(outputRow)
            End If
        End If
        If seg = "NK1" Then
            Dim outputRow As String() = New String() {id, "NK1", "Next of Kin", NK1.nk1Components(compIdx), NK1.nk1(queueIdx), convertedValue, DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ""}
            If parserInterface.showBlanksOnly = "yes" And NK1.nk1(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NK1.toBeParsedNk1.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NK1.toBeParsedNk1.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NK1.toBeParsedNk1.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (NK1.nk1(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NK1.toBeParsedNk1.Enqueue(outputRow)
            End If
        End If
        If seg = "PV1" Then
            Dim outputRow As String() = New String() {id, "PV1", "Patient Visit", PV1.pv1Components(compIdx), PV1.pv1(queueIdx), convertedValue, DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ""}
            If parserInterface.showBlanksOnly = "yes" And PV1.pv1(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PV1.toBeParsedPv1.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PV1.toBeParsedPv1.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PV1.toBeParsedPv1.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (PV1.pv1(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PV1.toBeParsedPv1.Enqueue(outputRow)
            End If
        End If
        If seg = "EVN" Then
            Dim outputRow As String() = New String() {id, "EVN", "Event Type", EVN.evnComponents(compIdx), EVN.evn(queueIdx), convertedValue, DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ""}
            If parserInterface.showBlanksOnly = "yes" And EVN.evn(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                EVN.toBeParsedEvn.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                EVN.toBeParsedEvn.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                EVN.toBeParsedEvn.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (EVN.evn(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                EVN.toBeParsedEvn.Enqueue(outputRow)
            End If
        End If
        If seg = "PID" Then
            Dim outputRow As String() = New String() {id, "PID", "Patient Identification", PID.pidComponents(compIdx), PID.pid(queueIdx), convertedValue, DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ""}
            If parserInterface.showBlanksOnly = "yes" And PID.pid(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PID.toBeParsedPid.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PID.toBeParsedPid.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PID.toBeParsedPid.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (PID.pid(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                PID.toBeParsedPid.Enqueue(outputRow)
            End If
        End If
        If seg = "OBR" Then
            Dim outputRow As String() = New String() {id, "OBR", "Observation Reports", OBR.obrComponents(compIdx), OBR.obr(queueIdx), convertedValue, DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ""}
            If parserInterface.showBlanksOnly = "yes" And OBR.obr(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBR.toBeParsedObr.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBR.toBeParsedObr.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBR.toBeParsedObr.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (OBR.obr(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBR.toBeParsedObr.Enqueue(outputRow)
            End If
        End If
        If seg = "OBX" Then
            Dim outputRow As String() = New String() {id, "OBX", "Observation", OBX.obxComponents(compIdx), OBX.obx(queueIdx), convertedValue, DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ""}
            If parserInterface.showBlanksOnly = "yes" And OBX.obx(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBX.toBeParsedObx.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBX.toBeParsedObx.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBX.toBeParsedObx.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (OBX.obx(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                OBX.toBeParsedObx.Enqueue(outputRow)
            End If
        End If
        If seg = "NTE" Then
            Dim outputRow As String() = New String() {id, "NTE", "Notes and Comments", NTE.nteComponents(compIdx), NTE.nte(queueIdx), convertedValue, DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), ""}
            If parserInterface.showBlanksOnly = "yes" And NTE.nte(queueIdx) = "" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NTE.toBeParsedNte.Enqueue(outputRow)
            ElseIf parserInterface.errorStatus = "yes" And outputRow(7) <> "" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NTE.toBeParsedNte.Enqueue(outputRow)
            ElseIf parserInterface.removeBlanks = False And parserInterface.errorStatus = "no" And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NTE.toBeParsedNte.Enqueue(outputRow)
            ElseIf ((parserInterface.removeBlanks = True) And (NTE.nte(queueIdx) <> "")) And parserInterface.showBlanksOnly = "no" Then
                parserInterface.DataGridView1.Rows.Add(outputRow)
                NTE.toBeParsedNte.Enqueue(outputRow)
            End If
        End If
    End Function
    Public Function comboBoxError(ByVal input As Queue(Of String()))
        For Each info In input
            If info(7) = "" Then
                Continue For
            Else
                parserInterface.DataGridView1.Rows.Add(info)
            End If
        Next
    End Function
    Public Function checkDTMrequired(ByVal id As Integer, ByVal input As String, ByVal seg As String, ByVal idx As Integer)
        If input = "" Then
            outputError(seg, id, idx, idx + 1, "Date/time of message is REQUIRED")
        ElseIf input.Length <= 2 Then
            outputError(seg, id, idx, idx + 1, "Invalid date/time of message")
        Else
            outputConverted(seg, id, idx, idx + 1, "")
        End If
    End Function
    Public Function checkDTMoptional(ByVal id As Integer, ByVal input As String, ByVal seg As String, ByVal idx As Integer)
        If input.Length = 1 Or input.Length = 2 Then
            outputError(seg, id, idx, idx + 1, "Invalid date/time of message")
        Else
            outputConverted(seg, id, idx, idx + 1, "")
        End If
    End Function
    Public Function usageRequired(ByVal id As Integer, ByVal input As String, ByVal seg As String, ByVal idx As Integer, ByVal msg As String)
        If input = "" Then
            outputError(seg, id, idx, idx + 1, msg)
        Else
            outputConverted(seg, id, idx, idx + 1, "")
        End If
    End Function
    Public Function dataTypeNM(ByVal id As Integer, ByVal input As String, ByVal seg As String, ByVal idx As Integer, ByVal msg As String)
        If IsNumeric(input) Or input = "" Then
            outputConverted(seg, id, idx, idx + 1, "")
        Else
            outputError(seg, id, idx, idx + 1, msg)
        End If
    End Function
    Public Function dataTypeSIoptional(ByVal id As Integer, ByVal input As String, ByVal seg As String, ByVal idx As Integer, ByVal msg As String)
        If (IsNumeric(input) And (input.Contains("-") = False) And (input.Contains(".") = False)) Or input = "" Then
            outputConverted(seg, id, idx, idx + 1, "")
        Else
            outputError(seg, id, idx, idx + 1, msg)
        End If
    End Function
    Public Function dataTypeSIrequired(ByVal id As Integer, ByVal input As String, ByVal seg As String, ByVal idx As Integer, ByVal msg As String)
        If IsNumeric(input) And (input.Contains("-") = False) And (input.Contains(".") = False) Then
            outputConverted(seg, id, idx, idx + 1, "")
        Else
            outputError(seg, id, idx, idx + 1, msg)
        End If
    End Function
End Module
