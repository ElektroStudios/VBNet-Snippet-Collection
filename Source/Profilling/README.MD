# Profilling snippets category
These snippets are intended to help manaing Regex aspects.

# General Information about this repository
 - .snippet files contains the formatted code that can be loaded through Visual Studio's code editor context menu.
 - .vb files contains the raw code that can be copied then pasted in any project.
 - Each .snippet and .vb file contains a #Region section and/or Xml documentation with code examples.
 
Feel free to use and/or modify any file of this repository.

If you like the job I've done, then please contribute with improvements to these snippets or by adding new ones.

# Profilling Util
Contains related profilling and unit testing utilities

This module was developed principally to validate and measure the code execution time of methods with the help of the 'ProfillingUtil.TestExecutionInfo' structure.

Public Members Summary

 - Types/Structures
   - ProfillingUtil.TestExecutionInfo

 - Constructors
   - ProfillingUtil.TestExecutionInfo.New(TimeSpan, MethodInfo, Boolean, Exception)

 - Properties
   - ProfillingUtil.Elapsed As TimeSpan
   - ProfillingUtil.Method As MethodInfo
   - ProfillingUtil.Success As Boolean
   - ProfillingUtil.Exception As Exception

 - Functions
   - ProfillingUtil.FlushMemory As Boolean
   - ProfillingUtil.FlushMemory(IntPtr) As Boolean
   - ProfillingUtil.TestSuccess(Action) As Boolean
   - ProfillingUtil.TestSuccessAsync(Action) As Task(Of Boolean)
   - ProfillingUtil.TestTime(Action) As TestExecutionInfo
   - ProfillingUtil.TestTimeAsync(Action) As Task(Of TestExecutionInfo)
   - ProfillingUtil.InlineAssignHelper(Of T)(T, T) As T
   - ProfillingUtil.InlineAssignHelper(Object, Object) As Object

 - Methods
   - ProfillingUtil.CollectGarbage
