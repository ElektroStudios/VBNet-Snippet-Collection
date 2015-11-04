
' -------------------
' Info of interest...
' -------------------
' 
' Operator Procedures: 
' http://msdn.microsoft.com/en-us/library/xh17yw4c.aspx
' 
' How to: Define an Operator: 
' http://msdn.microsoft.com/en-us/library/scs11cxx.aspx
' 
' How to: Define a Conversion Operator: 
' https://msdn.microsoft.com/en-us/library/yf7b9sy7.aspx
' 
' Operator Overloading: 
' http://msdn.microsoft.com/es-es/library/ms379613%28vs.80%29.aspx
' 
' Narrowing: 
' http://msdn.microsoft.com/en-us/library/4w127ed2.aspx
' 
' Widening: 
' http://msdn.microsoft.com/en-us/library/hz3wbx3x.aspx
' 
' Widening and Narrowing Conversions: 
' http://msdn.microsoft.com/en-us/library/k1e94s7e.aspx


<StructLayout(LayoutKind.Sequential)>
Public Structure ColorInfo

    Public Property R As Byte
    Public Property G As Byte
    Public Property B As Byte

    Public Sub New(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)

        Me.R = r
        Me.G = g
        Me.B = b

    End Sub

    Public Sub New(ByVal color As Color)

        Me.New(color.R, color.G, color.B)

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Performs an implicit conversion from <see cref="ColorInfo"/> to <see cref="Color"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="colorInfo">
    ''' The <see cref="colorInfo"/>.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The resulting <see cref="Color"/> of the conversion.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    Public Shared Widening Operator CType(ByVal colorInfo As ColorInfo) As Color

        Return Color.FromArgb(colorInfo.R, colorInfo.G, colorInfo.B)

    End Operator

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Performs an implicit conversion from <see cref="Color"/> to <see cref="ColorInfo"/>.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="color">
    ''' The <see cref="Color"/>.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The resulting <see cref="ColorInfo"/> of the conversion.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    Public Shared Narrowing Operator CType(ByVal color As Color) As ColorInfo

        Return New ColorInfo(color.R, color.G, color.B)

    End Operator

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Implements the operator =.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="colorInfo1">
    ''' The first <see cref="ColorInfo"/> to evaluate.
    ''' </param>
    ''' 
    ''' <param name="colorInfo2">
    ''' The second <see cref="ColorInfo"/> to evaluate.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The result of the operator.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    Public Shared Operator =(ByVal colorInfo1 As ColorInfo,
                             ByVal colorInfo2 As ColorInfo) As Boolean

        Return colorInfo1.Equals(colorInfo2)

    End Operator

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Implements the operator &lt;&gt;.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="colorInfo1">
    ''' The first <see cref="ColorInfo"/> to evaluate.
    ''' </param>
    ''' 
    ''' <param name="colorInfo2">
    ''' The second <see cref="ColorInfo"/> to evaluate.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The result of the operator.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    Public Shared Operator <>(ByVal colorInfo1 As ColorInfo,
                              ByVal colorInfo2 As ColorInfo) As Boolean

        Return Not colorInfo1.Equals(colorInfo2)

    End Operator

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Determines whether the specified <see cref="System.Object" /> is equal to this instance.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="obj">
    ''' Another object to compare to.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see langword="True"/> if the specified <see cref="System.Object" /> is equal to this instance; otherwise, <see langword="False"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    Public Overrides Function Equals(ByVal obj As Object) As Boolean

        If (TypeOf obj Is ColorInfo) Then
            Return Me.Equals(DirectCast(obj, ColorInfo))

        ElseIf (TypeOf obj Is Color) Then
            Return Me.Equals(New ColorInfo(DirectCast(obj, Color)))

        Else
            Return False

        End If

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Determines whether the specified <see cref="ColorInfo" /> is equal to this instance.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="colorInfo">
    ''' Another <see cref="ColorInfo"/> to compare to.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see langword="True"/> if the specified <see cref="ColorInfo" /> is equal to this instance; otherwise, <see langword="False"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    Public Overloads Function Equals(ByVal colorInfo As ColorInfo) As Boolean

        Return (colorInfo.R = Me.R) AndAlso (colorInfo.G = Me.G) AndAlso (colorInfo.B = Me.B)

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a hash code for this instance.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    Public Overrides Function GetHashCode() As Integer

        Return CType(Me, Color).GetHashCode()

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a <see cref="System.String" /> that represents this instance.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' A <see cref="System.String" /> that represents this instance.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    Public Overrides Function ToString() As String

        Return String.Format(CultureInfo.CurrentCulture, "{{R={0}, G={1}, B={2}}}", Me.R, Me.G, Me.B)

    End Function

End Structure
