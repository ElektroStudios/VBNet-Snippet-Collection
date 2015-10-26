# Generics snippets category
These snippets are intended to help managing .Net generics.

# General Information about this repository
 - .snippet files contains the formatted code that can be loaded through Visual Studio's code editor context menu.
 - .vb files contains the raw code that can be copied then pasted in any project.
 - Each .snippet and .vb file contains a #Region section and/or Xml documentation with code examples.
 
Feel free to use and/or modify any file of this repository.

If you like the job I've done, then please contribute with improvements to these snippets or by adding new ones.

# IEnumerable(Of T) Extensions
Contains custom extension methods to use with an IEnumerable(Of T)

Public Members Summary

 - Extension Functions
   - IEnumerable(Of T)().ConcatMultiple(IEnumerable(Of T)()) As IEnumerable(Of T)
   - IEnumerable(Of T)().StringJoin As IEnumerable(Of T)
   - IEnumerable(Of T).CountEmptyItems As Integer
   - IEnumerable(Of T).CountNonEmptyItems As Integer
   - IEnumerable(Of T).Duplicates As IEnumerable(Of T)
   - IEnumerable(Of T).Randomize As IEnumerable(Of T)
   - IEnumerable(Of T).RemoveDuplicates As IEnumerable(Of T)
   - IEnumerable(Of T).SplitIntoNumberOfElements(Integer) As IEnumerable(Of T)
   - IEnumerable(Of T).SplitIntoNumberOfElements(Integer, Boolean, T) As IEnumerable(Of T)
   - IEnumerable(Of T).SplitIntoParts(Integer) As IEnumerable(Of T)
   - IEnumerable(Of T).UniqueDuplicates As IEnumerable(Of T)
   - IEnumerable(Of T).Uniques As IEnumerable(Of T)

# IEnumerable(Of String) Extensions
Contains custom extension methods to use with an IEnumerable(Of String)

Public Members Summary

 - Extension Functions
   - IEnumerable(Of String).BubbleSort As IEnumerable(Of String)
   - IEnumerable(Of String).CountEmptyItems As Integer
   - IEnumerable(Of String).CountNonEmptyItems As Integer
   - IEnumerable(Of String).FindByContains(String, Boolean) As IEnumerable(Of String)
   - IEnumerable(Of String).FindByLike(String, Boolean) As IEnumerable(Of String)
   - IEnumerable(Of String).FindExact(String, StringComparison) As IEnumerable(Of String)
   - IEnumerable(Of String).RemoveByContains(String, Boolean) As IEnumerable(Of String)
   - IEnumerable(Of String).RemoveByLike(String, Boolean) As IEnumerable(Of String)
   - IEnumerable(Of String).RemoveExact(String, StringComparison) As IEnumerable(Of String)
