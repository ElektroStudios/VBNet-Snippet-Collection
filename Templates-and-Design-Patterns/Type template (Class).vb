
' -------------------
' Info of interest...
' -------------------
' 
' Structures and Classes:
' http://msdn.microsoft.com/en-us/library/2hkbth2a.aspx
' 
' Choosing Between Class and Struct:
' https://msdn.microsoft.com/en-us/library/ms229017%28v=vs.110%29.aspx
'
' Value Types and Reference Types:
' http://msdn.microsoft.com/en-us/library/t63sy5hs.aspx
'
' Walkthrough: Defining Classes:
' http://msdn.microsoft.com/en-us/library/xtka85tz.aspx
'
' Creating Classes:
' http://msdn.microsoft.com/en-us/library/ms973814.aspx
' 
' Using Classes and Structures.
' http://msdn.microsoft.com/en-us/library/aa289521%28v=vs.71%29.aspx

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Class description here.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Public NotInheritable Class MyType

#Region " Properties "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a <see cref="Object"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <value>
    ''' The <see cref="Object"/>.
    ''' </value>
    ''' ----------------------------------------------------------------------------------------------------
    Public ReadOnly Property MyProperty As Object
        Get
            Return Me.myPropertyB
        End Get
    End Property
    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' ( Backing field )
    ''' The <see cref="Object"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private ReadOnly myPropertyB As Object

#End Region

#Region " Constructors "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Prevents a default instance of the <see cref="MyType"/> class from being created.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerNonUserCode>
    Private Sub New()
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Initializes a new instance of the <see cref="MyType"/> class.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="value">
    ''' The value.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Sub New(ByVal value As Object)

        Me.myPropertyB = value

    End Sub

#End Region

End Class
