
' -------------------
' Info of interest...
' -------------------
' 
' Structures:
' http://msdn.microsoft.com/en-us/library/0awthy7k.aspx
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
' Structure Design:
' http://msdn.microsoft.com/en-us/library/vstudio/ms229031%28v=vs.100%29.aspx
'
' How to: Declare a Structure:
' http://msdn.microsoft.com/en-us/library/4ft0z102.aspx
'
' StructLayoutAttribute:
' http://msdn.microsoft.com/en-us/library/system.runtime.interopservices.structlayoutattribute.aspx
' 
' Using Classes and Structures.
' http://msdn.microsoft.com/en-us/library/aa289521%28v=vs.71%29.aspx


''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Structure description here.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
<StructLayout(LayoutKind.Sequential)>
Public Structure MyType

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
    ''' Initializes a new instance of the <see cref="MyType"/> structure.
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

End Structure
