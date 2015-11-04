
' -------------------
' Info of interest...
' -------------------
'
' Events:
' http://msdn.microsoft.com/en-us/library/ms172877.aspx
' 
' Event Handlers:
' http://msdn.microsoft.com/en-us/library/aa984105%28v=vs.71%29.aspx
'
' Handling and Raising Events:
' http://msdn.microsoft.com/en-us/library/edzehd2t%28v=vs.110%29.aspx
'
' AddHandler Statement:
' http://msdn.microsoft.com/en-us/library/7taxzxka.aspx
'
' RemoveHandler Statement:
' http://msdn.microsoft.com/en-us/library/3xz97kac.aspx
'
' EventHandler(Of TEventArgs) Delegate:
' http://msdn.microsoft.com/en-us/library/db0etb8x%28v=vs.110%29.aspx
' 
' Walkthrough: Declaring and Raising Events:
' http://msdn.microsoft.com/en-us/library/sc31b696.aspx
'
' How to: Create an Event and Handler:
' http://msdn.microsoft.com/en-us/library/1c6bkaht%28v=vs.90%29.aspx
' 
' How to: Raise and Consume Events:
' http://msdn.microsoft.com/en-us/library/9aackb16%28v=vs.110%29.aspx
' 
' How to: Declare Custom Events To Conserve Memory:
' http://msdn.microsoft.com/en-us/library/yt1k2w4e.aspx
' 
' How to: Declare Custom Events To Avoid Blocking (Visual Basic):
' http://msdn.microsoft.com/en-us/library/wf33s4w7.aspx


''' <summary>
''' Class Description.
''' </summary>
Public Class MyType

#Region " Events "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Occurs when the value changes.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public Event ValueChanged As EventHandler(Of ValueChangedEventArgs)

#End Region

#Region " Events Data "

#Region " MyEventArgs "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Contains the event-data of a <see cref="ValueChanged"/> event.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public NotInheritable Class ValueChangedEventArgs : Inherits EventArgs

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets my property.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' My property.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property MyProperty() As Integer
            Get
                Return Me.myPropertyB
            End Get
        End Property
        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' ( Backing field )
        ''' My property.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private ReadOnly myPropertyB As Integer

#End Region

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="MyEventArgs"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="MyEventArgs"/> class.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="value">
        ''' The value.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Public Sub New(ByVal value As Integer)

            Me.myPropertyB = value

        End Sub

#End Region

    End Class

#End Region

#End Region

#Region " Event Invocators "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Raises <see cref="ValueChanged"/> event.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="e">
    ''' The <see cref="ValueChangedEventArgs"/> instance containing the event data.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Protected Overridable Sub OnValueChanged(ByVal e As ValueChangedEventArgs)

        If (Me.ValueChangedEvent Is Nothing) Then
            MsgBox("nothing ")
            Exit Sub
        End If

        RaiseEvent ValueChanged(Me, e)

    End Sub

#End Region

#Region " Public Methods "

    Public Sub Test()

        For x As Integer = 0 To 10

            Me.OnValueChanged(New ValueChangedEventArgs(x))
        Next x

    End Sub

#End Region

End Class
