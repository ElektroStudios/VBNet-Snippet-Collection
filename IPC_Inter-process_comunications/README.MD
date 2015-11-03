# IPC (Inter-process communications) snippets category
These snippets are intended to help in IPC scenarios or also software automation.

# General Information about this repository
 - .snippet files contains the formatted code that can be loaded through Visual Studio's code editor context menu.
 - .vb files contains the raw code that can be copied then pasted in any project.
 - Each .snippet and .vb file contains a #Region section and/or Xml documentation with code examples.
 
Feel free to use and/or modify any file of this repository.

If you like the job I've done, then please contribute with improvements to these snippets or by adding new ones.

# IPC Util
Contains related Inter-process communication (IPC) utilities.

Public Members Summary

 - Child Classes
   - IpcUtil.SharedMemory
   - IpcUtil.UIAutomation

 - Functions
   - IpcUtil.SharedMemory.Create(String, Integer, Opt: MemoryMappedFileAccess) As MemoryMappedFile
   - IpcUtil.SharedMemory.Read(MemoryMappedFile, Long, Long) As Byte()
   - IpcUtil.SharedMemory.Read(String, Long, Long) As Byte()
   - IpcUtil.SharedMemory.ReadAt(MemoryMappedFile, Long) As Byte
   - IpcUtil.SharedMemory.ReadAt(String, Long) As Byte
   - IpcUtil.SharedMemory.ReadCharAt(MemoryMappedFile, Long) As Char
   - IpcUtil.SharedMemory.ReadCharAt(String, Long) As Char
   - IpcUtil.SharedMemory.ReadString(MemoryMappedFile, Long, Long, Opt: Encoding) As String
   - IpcUtil.SharedMemory.ReadString(String, Long, Long, Opt: Encoding) As String
   - IpcUtil.SharedMemory.ReadStringToEnd(MemoryMappedFile, Opt: Encoding) As String
   - IpcUtil.SharedMemory.ReadStringToEnd(String, Opt: Encoding) As String
   - IpcUtil.SharedMemory.ReadToEnd(MemoryMappedFile) As Byte()
   - IpcUtil.SharedMemory.ReadToEnd(String) As Byte()
   - IpcUtil.UIAutomation.GetTitlebarText(IntPtr) As String
   - IpcUtil.GetText(IntPtr) As String
   - IpcUtil.UIAutomation.SetText(IntPtr, String) As Boolean
   - IpcUtil.UIAutomation.AppendText(IntPtr, String) As Boolean
   - IpcUtil.UIAutomation.InsertText(IntPtr, Integer, String) As Boolean

 - Methods
   - IpcUtil.SharedMemory.Clear(MemoryMappedFile)
   - IpcUtil.SharedMemory.Clear(String)
   - IpcUtil.SharedMemory.Write(MemoryMappedFile, Byte())
   - IpcUtil.SharedMemory.Write(String, Byte())
   - IpcUtil.SharedMemory.WriteAt(MemoryMappedFile, Byte(), Long)
   - IpcUtil.SharedMemory.WriteAt(MemoryMappedFile, Byte, Long)
   - IpcUtil.SharedMemory.WriteAt(String, Byte(), Long)
   - IpcUtil.SharedMemory.WriteAt(String, Byte, Long)
   - IpcUtil.SharedMemory.WriteCharAt(MemoryMappedFile, Char, Long, Opt: Encoding)
   - IpcUtil.SharedMemory.WriteCharAt(String, Char, Long, Opt: Encoding)
   - IpcUtil.SharedMemory.WriteString(MemoryMappedFile, String, Opt: Encoding)
   - IpcUtil.SharedMemory.WriteString(String, String, Opt: Encoding)
   - IpcUtil.SharedMemory.WriteStringAt(MemoryMappedFile, String, Long, Opt: Encoding)
   - IpcUtil.SharedMemory.WriteStringAt(String, String, Long, Opt: Encoding)

# Sharing memory across different running applications example
Example for beginners to share a memory segment between two applications.
