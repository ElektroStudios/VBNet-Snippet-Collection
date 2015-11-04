# Process snippets category
These snippets are intended to help with process management.

# General Information about this repository
 - .snippet files contains the formatted code that can be loaded through Visual Studio's code editor context menu.
 - .vb files contains the raw code that can be copied then pasted in any project.
 - Each .snippet and .vb file contains a #Region section and/or Xml documentation with code examples.
 
Feel free to use and/or modify any file of this repository.

If you like the job I've done, then please contribute with improvements to these snippets or by adding new ones.

# Process Util
Contains related process utilities, such a process-execution monitor.

Public Members Summary

 - Types
   - ProcessUtil.ProcessMonitor : Implements IDisposable
   - ProcessUtil.ProcessMonitor.ProcessMonitorEventArgs : Inherits EventArgs
   - ProcessUtil.ProcessWatcher

 - Constructors
   - ProcessUtil.ProcessMonitor.New()
   - ProcessUtil.ProcessMonitor.ProcessMonitorEventArgs.New(Integer, String)
   - ProcessUtil.ProcessWatcher.New()

 - Events
   - ProcessUtil.ProcessMonitor.ProcessStarted As EventHandler(Of ProcessUtil.ProcessMonitor.ProcessMonitorEventArgs)
   - ProcessUtil.ProcessMonitor.ProcessStopped As EventHandler(Of ProcessUtil.ProcessMonitor.ProcessMonitorEventArgs)
   - ProcessUtil.ProcessWatcher.ProcessStarted As EventArrivedEventHandler
   - ProcessUtil.ProcessWatcher.ProcessStopped As EventArrivedEventHandler

 - Properties
   - ProcessUtil.ProcessMonitor.Interval As Integer
   - ProcessUtil.ProcessMonitor.ProcessMonitorEventArgs.ProcessId As Integer
   - ProcessUtil.ProcessMonitor.ProcessMonitorEventArgs.ProcessName As String

 - Fucntions
   - ProcessUtil.CloseProcess(String, Opt: Boolean) As Boolean
   - ProcessUtil.CloseProcesses(String, Opt: Boolean) As Boolean
   - ProcessUtil.GetProcessesValue(Of T)(String, Func(Of Process, T), Opt: Boolean) As IEnumerable(Of T)
   - ProcessUtil.GetProcessValue(Of T)(String, Func(Of Process, T), Opt: Boolean) As T
   - ProcessUtil.GetWmiProcessQuery(KeyValuePair(Of String, String)) As ManagementObject()
   - ProcessUtil.IsRunning(String) As Boolean
   - ProcessUtil.KillProcess(String, Opt: Boolean, Opt: Boolean) As Boolean
   - ProcessUtil.KillProcesses(String, Opt: Boolean, Opt: Boolean) As Boolean
   - ProcessUtil.SetProcessesValue(String, Action(Of Process), Opt: Boolean)
   - ProcessUtil.SetProcessesValues(String, Action(Of Process), Opt: Boolean)

 - Methods
   - ProcessUtil.WaitUntilLoaded(Process, Opt: Integer)
   - ProcessUtil.WaitUntilLoadedAsync(Process, Action, Opt: Integer)
   - ProcessUtil.ProcessMonitor.Start
   - ProcessUtil.ProcessMonitor.Stop
   - ProcessUtil.ProcessMonitor.Dispose
   - ProcessUtil.ProcessWatcher.Start
   - ProcessUtil.ProcessWatcher.Stop

# Read Process StdOut-StdErr Asynchronouslly example
Example for beginners that demonstrates how to read StdOut-StdErr outputs of a pocess, asynchronouslly.

# Read Process StdOut-StdErr Synchronouslly example
Example that demonstrates how to read StdOut-StdErr outputs of a pocess, synchronouslly.
