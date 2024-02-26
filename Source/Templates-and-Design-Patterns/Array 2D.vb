' -------------------
' Info of interest...
' -------------------
' 
' Overview of Arrays:
' http://msdn.microsoft.com/en-us/library/2k7ayc03%28v=vs.90%29.aspx
' 
' Array Dimensions:
' http://msdn.microsoft.com/en-us/library/02e7z943.aspx
' 
' Multidimensional Arrays:
' http://msdn.microsoft.com/en-us/library/d2de1t93%28v=vs.90%29.aspx
'
' How to: Initialize a Multidimensional Array:
' http://msdn.microsoft.com/en-us/library/0sxy840k%28v=vs.90%29.aspx


' Create a multidimensional Array of 2x2 dimensions.
Dim matrix As String(,) = New String(2, 1) {
    {"Item 0,0", "Item 0,1"},
    {"Item 1,0", "Item 1,1"},
    {"Item 2,0", "Item 2,1"}
}

' Set a value.
matrix(0, 1) = "New Item 0,1"

' Get a Value.
MessageBox.Show(matrix(0, 1))

' Loop through the Array bounds.
For iOuter As Integer = matrix.GetLowerBound(0) To matrix.GetUpperBound(0) ' iOuter represents the first dimension.

    For iInner As Integer = matrix.GetLowerBound(1) To matrix.GetUpperBound(1) ' iInner represents the second dimension.

        Console.WriteLine(String.Format("Array 2D {0},{1}: {2}", iOuter, iInner, matrix(iOuter, iInner)))

    Next iInner

Next iOuter
