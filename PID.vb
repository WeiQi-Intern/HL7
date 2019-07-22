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


        For idx As Integer = 0 To 29
            Dim outputRow As String() = New String() {id, "PID", "Patient Identification", pidComponents(idx), pid(idx + 1), "NULL", DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "NULL"}
            parserInterface.DataGridView1.Rows.Add(outputRow)
            toBeParsedPid.Enqueue(outputRow)
        Next
    End Function
End Class
