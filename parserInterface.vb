Public Class parserInterface
    Public timerCounter As Integer = 100
    Public existingFiles As Queue(Of String) = New Queue(Of String)
    Public parseContent As List(Of String) = New List(Of String)
    Public parseIndex As List(Of Integer) = New List(Of Integer)
    Public piMsg As String() = {}
    Public piIndex As Integer = 0
    Public removeBlanks As Boolean = False
    Public errorStatus As String = "no"
    Public showBlanksOnly As String = "no"
    Public Function importHL7()
        If System.IO.Directory.Exists("C:\Users\weiqi-intern\Desktop\HL7-samples") Then
            Dim dir As New System.IO.DirectoryInfo("C:\Users\weiqi-intern\Desktop\HL7-samples")
            Dim fileArr As System.IO.FileInfo() = dir.GetFiles()
            Dim numSamples As Integer = fileArr.Length
            Dim id As Integer = 0
            For n As Integer = 0 To numSamples - 1
                Dim curr As String = fileArr(n).FullName
                Dim content = My.Computer.FileSystem.ReadAllText(curr)
                Dim row As String()
                If ((fileArr(n).Name = "desktop.ini") Or (existingFiles.Contains(fileArr(n).Name)) = True) Then
                    Continue For
                Else
                    id = existingFiles.Count + 1
                    row = New String() {id, content, DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), fileArr(n).Name, fileArr(n).FullName, ""}
                    TextBox2.Text = DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
                    existingFiles.Enqueue(fileArr(n).Name)
                    'will call these during parsing (used in form parse.vb)
                    parseContent.Add(content)
                    parseIndex.Add(id)
                End If
                DataGridView2.Rows.Add(row)
            Next
            TextBox4.Visible = True
        Else
            TextBox4.Text = DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "  Error: Directory C:\Users\weiqi-intern\Desktop\HL7-samples is not found!"
            TextBox4.Visible = True
        End If
        Timer1.Enabled = True
        Return True
    End Function
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        timerCounter -= 1
        If timerCounter = 0 Then
            Timer1.Enabled = False
            'example of 100s countdown, auto update every 100s
            timerCounter = 100
            importHL7()
        End If
        TextBox3.Text = timerCounter

        'row turns peachpuff when there is error!
        For i As Integer = 0 To DataGridView1.Rows.Count() - 1 Step +1
            Dim val As String
            val = DataGridView1.Rows(i).Cells(7).Value
            If val <> "" Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.PeachPuff
            End If
        Next
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        importHL7()
        TextBox4.Visible = True
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        parseContent.Clear()
        existingFiles.Clear()
        DataGridView2.Rows.Clear()
        TextBox1.Text = DateAndTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
        DataGridView1.Rows.Clear()
        DataGridView1.Visible = False
        Label2.Visible = False
        ComboBox1.Visible = False
        TextBox4.Visible = False
        TextBox5.Visible = False
    End Sub
    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        'rmb clearr!! here
        'MSH.msh.Clear()
        removeBlanks = False
        errorStatus = "no"
        showBlanksOnly = "no"
        DataGridView1.Rows.Clear()
        ComboBox1.ResetText()
        TextBox4.Clear()

        Dim index As Integer
        index = e.RowIndex
        Dim selectedRow As DataGridViewRow
        selectedRow = DataGridView2.Rows(index)

        'message array where each element is one segment
        Dim rawMessage As String = ""
        Dim id As Integer = 0
        If parseContent.Count <> 0 Then
            rawMessage = parseContent(index)
            id = parseIndex(index)
            'split into different segment in the message array below
            Dim message As String() = rawMessage.Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
            parsing.startParser(message, id)
            piMsg = message
            piIndex = id
            DataGridView1.Visible = True
            Label2.Visible = True
            ComboBox1.Visible = True
            TextBox5.Visible = True
        Else
            MessageBox.Show("Please wait while the system loads...")
            Label2.Visible = False
            DataGridView1.Visible = False
            ComboBox1.Visible = False
            TextBox5.Visible = False
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'MSH.msh.Clear()
        removeBlanks = False
        errorStatus = "no"
        showBlanksOnly = "no"
        DataGridView1.Rows.Clear()
        If ComboBox1.Text = "All" Or ComboBox1.Text = "" Then
            removeBlanks = False
            parsing.startParser(piMsg, piIndex)
        ElseIf ComboBox1.Text = "Show Filled Components Only" Then
            removeBlanks = True
            parsing.startParser(piMsg, piIndex)
        ElseIf ComboBox1.Text = "Show Unfilled Components Only" Then
            showBlanksOnly = "yes"
            parsing.startParser(piMsg, piIndex)
        ElseIf ComboBox1.Text = "MSH – Message Header" Then
            For Each piSeg In piMsg
                If (piSeg(0) + piSeg(1) + piSeg(2)) = "MSH" Then
                    MSH.parseMSH(piSeg, piIndex)
                Else
                    Continue For
                End If
            Next
        ElseIf ComboBox1.Text = "MSA – Message Acknowledgment" Then
            For Each piSeg In piMsg
                If (piSeg(0) + piSeg(1) + piSeg(2)) = "MSA" Then
                    MSA.parseMSA(piSeg, piIndex)
                Else
                    Continue For
                End If
            Next
        ElseIf ComboBox1.Text = "EVN – Event Type" Then
            For Each piSeg In piMsg
                If (piSeg(0) + piSeg(1) + piSeg(2)) = "EVN" Then
                    EVN.parseEVN(piSeg, piIndex)
                Else
                    Continue For
                End If
            Next
        ElseIf ComboBox1.Text = "PID – Patient Identification" Then
            For Each piSeg In piMsg
                If (piSeg(0) + piSeg(1) + piSeg(2)) = "PID" Then
                    PID.parsePID(piSeg, piIndex)
                Else
                    Continue For
                End If
            Next
        ElseIf ComboBox1.Text = "NK1 – Next of Kin/Associated Parties" Then
            For Each piSeg In piMsg
                If (piSeg(0) + piSeg(1) + piSeg(2)) = "NK1" Then
                    NK1.parseNK1(piSeg, piIndex)
                Else
                    Continue For
                End If
            Next
        ElseIf ComboBox1.Text = "PV1 – Patient Visit" Then
            For Each piSeg In piMsg
                If (piSeg(0) + piSeg(1) + piSeg(2)) = "PV1" Then
                    PV1.parsePV1(piSeg, piIndex)
                Else
                    Continue For
                End If
            Next
        ElseIf ComboBox1.Text = "OBR – Observation Reports" Then
            For Each piSeg In piMsg
                If (piSeg(0) + piSeg(1) + piSeg(2)) = "OBR" Then
                    OBR.parseOBR(piSeg, piIndex)
                Else
                    Continue For
                End If
            Next
        ElseIf ComboBox1.Text = "OBX – Observation" Then
            For Each piSeg In piMsg
                If (piSeg(0) + piSeg(1) + piSeg(2)) = "OBX" Then
                    OBX.parseOBX(piSeg, piIndex)
                Else
                    Continue For
                End If
            Next
        ElseIf ComboBox1.Text = "NTE – Notes and Comments (Observation Specific)" Then
            For Each piSeg In piMsg
                If (piSeg(0) + piSeg(1) + piSeg(2)) = "NTE" Then
                    NTE.parseNTE(piSeg, piIndex)
                Else
                    Continue For
                End If
            Next
        ElseIf ComboBox1.Text = "Show Errors Only" Then
            removeBlanks = False
            errorStatus = "yes"
            parsing.startParser(piMsg, piIndex)
        End If
    End Sub
End Class