
' -------------------
' Info of interest...
' -------------------
' 
' Process.OutputDataReceived Event:
' http://msdn.microsoft.com/en-us/library/system.diagnostics.process.outputdatareceived%28v=vs.110%29.aspx
'
' Process.ErrorDataReceived Event:
' http://msdn.microsoft.com/en-us/library/system.diagnostics.process.errordatareceived%28v=vs.110%29.aspx
' 
' DataReceivedEventHandler Delegate:
' http://msdn.microsoft.com/en-us/library/system.diagnostics.datareceivedeventhandler%28v=vs.110%29.aspx


''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' The process to execute.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Private WithEvents proc As New Process With
    {
        .EnableRaisingEvents = True,
        .StartInfo = New ProcessStartInfo With
                     {
                         .FileName = "CMD.exe",
                         .Arguments = "/C Dir /B ""C:\Windows\System32""",
                         .WorkingDirectory = Application.StartupPath,
                         .WindowStyle = ProcessWindowStyle.Hidden,
                         .UseShellExecute = False,
                         .CreateNoWindow = True,
                         .RedirectStandardError = True,
                         .RedirectStandardOutput = True,
                         .StandardErrorEncoding = Encoding.Default,
                         .StandardOutputEncoding = Encoding.Default
                     }
    }

Sub Test()

    With Me.proc
        .Start()
        .BeginOutputReadLine()
        .BeginErrorReadLine()
        ' .WaitForExit(milliseconds:=0)
        .WaitForExit()
    End With

End Sub

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Handles the <see cref="Process.OutputDataReceived"/> event of the <see cref="proc"/> instance.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
''' <param name="sender">
''' The source of the event.
''' </param>
''' 
''' <param name="e">
''' The <see cref="DataReceivedEventArgs"/> instance containing the event data.
''' </param>
''' ----------------------------------------------------------------------------------------------------
Private Sub Process_OutputDataReceived(ByVal sender As Object, ByVal e As DataReceivedEventArgs) _
Handles Process.OutputDataReceived

    If Not String.IsNullOrEmpty(e.Data) Then
        Console.WriteLine(String.Format("stdOut: {0}", e.Data))
    End If

End Sub

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Handles the <see cref="Process.ErrorDataReceived"/> event of the <see cref="proc"/> instance.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
''' <param name="sender">The source of the event.
''' </param>
''' 
''' <param name="e">
''' The <see cref="DataReceivedEventArgs"/> instance containing the event data.
''' </param>
''' ----------------------------------------------------------------------------------------------------
Private Sub Process_ErrorDataReceived(ByVal sender As Object, ByVal e As DataReceivedEventArgs) _
Handles Process.ErrorDataReceived

    If Not String.IsNullOrEmpty(e.Data) Then
        Console.WriteLine(String.Format("stdErr: {0}", e.Data))
    End If

End Sub

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Handles the <see cref="Process.Exited"/> event of the <see cref="proc"/> instance.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
''' <param name="sender">
''' The source of the event.
''' </param>
''' 
''' <param name="e">
''' The <see cref="EventArgs"/> instance containing the event data.
''' </param>
''' ----------------------------------------------------------------------------------------------------
Private Sub Process_Exited(ByVal sender As Object, ByVal e As EventArgs) _
Handles Process.Exited

    Console.WriteLine(String.Format("Process exited at {0}", Date.Now.ToShortTimeString))

End Sub
