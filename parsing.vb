Public Class parsing
    Public Shared Function parseSegment(ByVal seg As String, ByVal line As String, ByVal id As Integer)
        If seg = "MSH" Then
            MSH.parseMSH(line, id)
        End If
        If seg = "MSA" Then
            MSA.parseMSA(line, id)
        End If
        If seg = "EVN" Then
            EVN.parseEVN(line, id)
        End If
        'here onwards, segment can be multiple
        If seg = "PID" Then
            PID.parsePID(line, id)
        End If
        If seg = "NK1" Then
            NK1.parseNK1(line, id)
        End If
        If seg = "PV1" Then
            PV1.parsePV1(line, id)
        End If
        If seg = "OBR" Then
            OBR.parseOBR(line, id)
        End If
        If seg = "OBX" Then
            OBX.parseOBX(line, id)
        End If
        If seg = "NTE" Then
            NTE.parseNTE(line, id)
        End If
    End Function
    Public Shared Function startParser(ByVal input As String(), ByVal id As Integer)
        'check if HL7 messsage is valid
        If ((input.Count = 0) OrElse (input(0)(0) + input(0)(1) + input(0)(2) <> "MSH")) Then
            parserInterface.TextBox4.Text = "Invalid HL7 message! Should begin with MSH."
            parserInterface.TextBox4.Visible = True
            parserInterface.DataGridView1.Visible = False
        Else
            For Each line In input
                Dim seg As String = line(0) + line(1) + line(2)
                If ((seg = "MSH") Or (seg = "MSA") Or (seg = "EVN") Or (seg = "PID") Or (seg = "NK1") Or (seg = "PV1") Or (seg = "OBR") Or (seg = "OBX") Or (seg = "NTE")) Then
                    parseSegment(seg, line, id)
                End If
            Next
        End If
    End Function
End Class
