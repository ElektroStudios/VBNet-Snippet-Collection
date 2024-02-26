
' -------------------
' Info of interest...
' -------------------
' 
' Serialization in the .NET Framework:
' http://msdn.microsoft.com/en-us/library/7ay27kt9%28v=vs.110%29.aspx
' 
' Implement ISerializable correctly:
' http://msdn.microsoft.com/en-us/library/ms182342.aspx
'
' ISerializable Interface
' http://msdn.microsoft.com/en-us/library/system.runtime.serialization.iserializable%28v=vs.110%29.aspx
'
' IXmlSerializable Interface
' http://msdn.microsoft.com/en-us/library/System.Xml.Serialization.IXmlSerializable%28v=vs.110%29.aspx


#Region " Imports "

Imports System.Runtime.Serialization
Imports System.Security.Permissions
Imports System.Xml
Imports System.Xml.Serialization

#End Region

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Class description.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
''' <remarks>
''' This class can be serialized.
''' </remarks>
''' ----------------------------------------------------------------------------------------------------
<Serializable>
<XmlRoot("MySerializableClass_RootName")>
Public NotInheritable Class MySerializableClass : Implements ISerializable, IXmlSerializable

#Region " Properties "

    Public Property StrValue As String
    Public Property Int32Value As Integer

#End Region

#Region " Constructors "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Prevents a default instance of the <see cref="MySerializableClass"/> class from being created.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Private Sub New()
    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Initializes a new instance of the <see cref="MySerializableClass"/> class.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="strValue">
    ''' A <see cref="String"/> value.
    ''' </param>
    ''' 
    ''' <param name="int32Value">
    ''' A <see cref="Integer"/> value.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Public Sub New(ByVal strValue As String,
                   ByVal int32Value As Integer)

        Me.StrValue = strValue
        Me.Int32Value = int32Value

    End Sub

#End Region

#Region " ISerializable implementation " ' For Binary serialization.

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Populates a <see cref="T:SerializationInfo"/> with the data needed to serialize the target object.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="info">
    ''' The <see cref="T:SerializationInfo"/> to populate with data.
    ''' </param>
    ''' 
    ''' <param name="context">
    ''' The destination (see <see cref="T:StreamingContext"/>) for this serialization.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <exception cref="ArgumentNullException">
    ''' info
    ''' </exception>
    ''' ----------------------------------------------------------------------------------------------------
    <SecurityPermissionAttribute(SecurityAction.LinkDemand, Flags:=SecurityPermissionFlag.SerializationFormatter)>
    Protected Sub GetObjectData(ByVal info As SerializationInfo, ByVal context As StreamingContext) Implements ISerializable.GetObjectData

        If info Is Nothing Then
            Throw New ArgumentNullException(paramName:="info")
        End If

        With info
            .AddValue("PropertyName1", Me.StrValue, Me.StrValue.GetType)
            .AddValue("PropertyName2", Me.Int32Value, Me.Int32Value.GetType)
        End With

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Initializes a new instance of the <see cref="MySerializableClass"/> class.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' This constructor is used to deserialize values.
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="info">
    ''' The <see cref="T:SerializationInfo"/> to populate with data.
    ''' </param>
    ''' 
    ''' <param name="context">
    ''' The destination (see <see cref="T:StreamingContext"/>) for this deserialization.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <exception cref="ArgumentNullException">
    ''' info
    ''' </exception>
    ''' ----------------------------------------------------------------------------------------------------
    Protected Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)

        If info Is Nothing Then
            Throw New ArgumentNullException("info")
        End If

        Me.StrValue = info.GetString("PropertyName1")
        Me.Int32Value = info.GetInt32("PropertyName2")

    End Sub

#End Region

#Region " IXMLSerializable implementation " ' For Xml serialization.

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' This method is reserved and should not be used.
    ''' When implementing the <see cref="IXmlSerializable"/> interface, you should return <see langword="Nothing"/> from this method, 
    ''' and instead, if specifying a custom schema is required, apply the <see cref="T:XmlSchemaProviderAttribute"/> to the class.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' An <see cref="T:Xml.Schema.XmlSchema"/> that describes the Xml representation of the object 
    ''' that is produced by the <see cref="M:IXmlSerializable.WriteXml(Xml.XmlWriter)"/> method 
    ''' and consumed by the <see cref="M:IXmlSerializable.ReadXml(Xml.XmlReader)"/> method.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    Public Function GetSchema() As Schema.XmlSchema Implements IXmlSerializable.GetSchema

        Return Nothing

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Converts an object into its Xml representation.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="writer">
    ''' The <see cref="T:Xml.XmlWriter"/> stream to which the object is serialized.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Public Sub WriteXml(ByVal writer As XmlWriter) Implements IXmlSerializable.WriteXml

        writer.WriteElementString("PropertyName1", Me.StrValue)
        writer.WriteElementString("PropertyName2", Me.Int32Value.ToString)

    End Sub

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Generates an object from its Xml representation.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="reader">
    ''' The <see cref="T:Xml.XmlReader"/> stream from which the object is deserialized.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    Public Sub ReadXml(ByVal reader As XmlReader) Implements IXmlSerializable.ReadXml

        With reader

            .ReadStartElement(MyBase.GetType.Name)

            Me.StrValue = .ReadElementContentAsString
            Me.Int32Value = .ReadElementContentAsInt

        End With

    End Sub

#End Region

End Class
