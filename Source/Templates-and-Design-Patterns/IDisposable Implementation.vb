
' -------------------
' Info of interest...
' -------------------
' 
' IDisposable interface:
' http://msdn.microsoft.com/en-us/library/system.idisposable%28v=vs.110%29.aspx
'
' Implementing a Dispose Method:
' http://msdn.microsoft.com/en-us/library/fs2xkftw%28v=vs.110%29.aspx
'
' Dispose Pattern:
' http://msdn.microsoft.com/en-us/library/b1yfkh5e%28v=vs.110%29.aspx
'
' Object.Finalize Method:
' http://msdn.microsoft.com/en-us/library/system.object.finalize%28v=vs.110%29.aspx
' 
' GC.SuppressFinalize Method:
' http://msdn.microsoft.com/en-us/library/system.gc.suppressfinalize.aspx
'
' Object Lifetime: How Objects Are Created and Destroyed:
' http://msdn.microsoft.com/en-us/library/hks5e2k6.aspx
' 
' Fundamentals of Garbage Collection:
' http://msdn.microsoft.com/en-us/library/ee787088%28v=vs.110%29.aspx
'
' Cleaning Up Unmanaged Resources:
' http://msdn.microsoft.com/en-us/library/498928w2.aspx


Public Class DisposableType : Implements IDisposable

    Public Sub TestMethod()

        ' Check if this instance is disposed.
        Me.DisposedCheck()
        ' If is not disposed then continue.
        ' Do something...

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' To detect redundant calls when disposing.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private isDisposed As Boolean

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Prevent calls to methods after disposing.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <exception cref="System.ObjectDisposedException">
    ''' </exception>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Private Sub DisposedCheck()

        If (Me.isDisposed) Then
            Throw New ObjectDisposedException(Me.GetType().FullName)
        End If

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Releases all the resources used by this instance.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Sub Dispose() Implements IDisposable.Dispose

        Me.Dispose(isDisposing:=True)
        GC.SuppressFinalize(obj:=Me)

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
    ''' Releases unmanaged and - optionally - managed resources.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="isDisposing">
    ''' <see langword="True"/>  to release both managed and unmanaged resources; 
    ''' <see langword="False"/> to release only unmanaged resources.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Protected Overridable Sub Dispose(ByVal isDisposing As Boolean)

        If (Not Me.isDisposed) AndAlso (isDisposing) Then

            ' Dispose managed objects here...

        End If

        Me.isDisposed = True

    End Sub

End Class
