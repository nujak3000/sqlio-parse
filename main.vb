Module Module1
 
    Public CurrentDir As String = System.Environment.CurrentDirectory &amp; ""
 
    Public OutputFile As String = CurrentDir &amp; Format$(Date.Now, "yyyyMMddHHmmss") &amp; "output.csv"
 
    Public TempDir As String = CurrentDir &amp; "Results"
 
    Public dir_info As New IO.DirectoryInfo(TempDir)
 
    Sub Main()
        parseresults()
    End Sub
 
    Function parseresults() As Boolean
        Try
            Dim file_info_ar As IO.FileInfo() = dir_info.GetFiles()
            Dim file_info As IO.FileInfo
            Dim line As String
            Dim contents() As String
            Dim itemcounter As Integer = 0
            Dim csvfile As String = "time,cmds,iops,bw,late" &amp; System.Environment.NewLine
 
            For Each file_info In file_info_ar
                contents = System.IO.File.ReadAllLines(TempDir &amp; file_info.Name)
                For Each line In contents
                    If line.Contains("C:\SQLIO>sqlio") Then
                        csvfile = csvfile & file_info.Name &amp; "," &amp; line
                    End If
                    If line.Contains("IOs/sec:") Then
                        csvfile = csvfile & "," &amp; line.Substring(8, line.Length - 8)
                    End If
                    If line.Contains("MBs/sec:") Then
                        csvfile = csvfile & "," &amp; line.Substring(8, line.Length - 8)
                    End If
                    If line.Contains("Avg_Latency") Then
                        csvfile = csvfile & "," &amp; line.Substring(17, line.Length - 17) & System.Environment.NewLine
                    End If
                Next
            Next
 
            System.IO.File.WriteAllText(OutputFile, csvfile)
 
        Catch ex As Exception
            System.IO.File.AppendAllText("exceptions.log", ex.ToString)
            Return False
        End Try
        Return True
    End Function
End Module
