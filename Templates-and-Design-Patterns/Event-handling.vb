
' -------------------
' Info of interest...
' -------------------
'
' Events:
' http://msdn.microsoft.com/en-us/library/ms172877.aspx
' 
' Handles Clause:
' http://msdn.microsoft.com/en-us/library/6k46st1y.aspx
'
' AddHandler Statement:
' http://msdn.microsoft.com/en-us/library/7taxzxka.aspx
'
' RemoveHandler Statement:
' http://msdn.microsoft.com/en-us/library/3xz97kac.aspx
'
' EventHandler(Of TEventArgs) Delegate:
' http://msdn.microsoft.com/en-us/library/db0etb8x%28v=vs.110%29.aspx


Public Class Form1 : Inherits Form

    Friend WithEvents TextBox1 As New TextBox
    Friend WithEvents TextBox2 As New TextBox

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.TextBox1 = New TextBox
        Me.TextBox2 = New TextBox

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the <see cref="TextBox.TextChanged"/> event of the <see cref="TextBox1"/> and <see cref="TextBox2"/> controls.
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
    Private Sub TextBoxes_TextChanged(ByVal sender As Object, ByVal e As EventArgs) _
    Handles TextBox1.TextChanged,
            TextBox2.TextChanged

        Dim tb As TextBox = DirectCast(sender, TextBox)

        ' Your code goes here...

    End Sub

End Class
