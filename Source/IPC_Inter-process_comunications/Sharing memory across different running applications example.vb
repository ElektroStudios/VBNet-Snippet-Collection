' INSTRUCTIONS:
' 1. Create a new project, copy&paste the 'MemoryMappedFile_Form1' Form and compile it.
' 2. Create a new project, copy&paste the 'MemoryMappedFile_Form2' Form and compile it.
' 3. Run both applications separately to test their memory messages.



' *************************
' This is the Application 1
' *************************

#Region " Imports "

Imports System.IO.MemoryMappedFiles

#End Region

#Region " Application 1 "

''' <summary>
''' Class MemoryMappedFile_Form1.
''' This should be the Class used to compile our first application.
''' </summary>
Public Class MemoryMappedFile_Form1

    ' The controls to create on execution-time.
    Dim WithEvents btMakeFile As New Button ' Writes the memory.
    Dim WithEvents btReadFile As New Button ' Reads the memory.
    Dim tbMessage As New TextBox ' Determines the string to map into memory.
    Dim tbReceptor As New TextBox ' Print the memory read's result.
    Dim lbInfoButtons As New Label ' Informs the user with a usage hint for the buttons.
    Dim lbInfotbMessage As New Label ' Informs the user with a usage hint for 'tbMessage'.

    ''' <summary>
    ''' Indicates the name of our memory-file.
    ''' </summary>
    Private ReadOnly MemoryName As String = "My Memory-File Name"

    ''' <summary>
    ''' Indicates the memory buffersize to store the <see cref="MemoryName"/>, in bytes.
    ''' </summary>
    Private ReadOnly MemoryBufferSize As Integer = 1024I

    ''' <summary>
    ''' Indicates the string to map in memory.
    ''' </summary>
    Private ReadOnly Property strMessage As String
        Get
            Return tbMessage.Text
        End Get
    End Property

    ''' <summary>
    ''' Initializes a new instance of the <see cref="MemoryMappedFile_Form1"/> class.
    ''' </summary>
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Set the properties of the controls.
        With lbInfotbMessage
            .Location = New Point(20, 10)
            .Text = "Type in this TextBox the message to write in memory:"
            .AutoSize = True
            ' .Size = tbReceptor.Size
        End With
        With tbMessage
            .Text = "Hello world from application one!"
            .Location = New Point(20, 30)
            .Size = New Size(310, Me.tbMessage.Height)
        End With
        With btMakeFile
            .Text = "Write Memory"
            .Size = New Size(130, 45)
            .Location = New Point(20, 50)
        End With
        With btReadFile
            .Text = "Read Memory"
            .Size = New Size(130, 45)
            .Location = New Point(200, 50)
        End With
        With tbReceptor
            .Location = New Point(20, 130)
            .Size = New Size(310, 100)
            .Multiline = True
        End With
        With lbInfoButtons
            .Location = New Point(tbReceptor.Location.X, tbReceptor.Location.Y - 30)
            .Text = "Press '" & btMakeFile.Text & "' button to create the memory file, that memory can be read from both applications."
            .AutoSize = False
            .Size = tbReceptor.Size
        End With

        ' Set the Form properties.
        With Me
            .Text = "Application 1"
            .Size = New Size(365, 300)
            .FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
            .MaximizeBox = False
            .StartPosition = FormStartPosition.CenterScreen
        End With

        ' Add the controls on the UI.
        Me.Controls.AddRange({lbInfotbMessage, tbMessage, btMakeFile, btReadFile, tbReceptor, lbInfoButtons})

    End Sub

    ''' <summary>
    ''' Writes a byte sequence into a <see cref="MemoryMappedFile"/>.
    ''' </summary>
    ''' <param name="Name">Indicates the name to assign the <see cref="MemoryMappedFile"/>.</param>
    ''' <param name="BufferLength">Indicates the <see cref="MemoryMappedFile"/> buffer-length to write in.</param>
    ''' <param name="Data">Indicates the byte-data to write inside the <see cref="MemoryMappedFile"/>.</param>
    Private Sub MakeMemoryMappedFile(ByVal Name As String, ByVal BufferLength As Integer, ByVal Data As Byte())

        ' Create or open the memory-mapped file.
        Dim MessageFile As MemoryMappedFile =
            MemoryMappedFile.CreateOrOpen(Name, Me.MemoryBufferSize, MemoryMappedFileAccess.ReadWrite)

        ' Write the byte-sequence into memory.
        Using Writer As MemoryMappedViewAccessor =
            MessageFile.CreateViewAccessor(0L, Me.MemoryBufferSize, MemoryMappedFileAccess.ReadWrite)

            ' Firstly fill with null all the buffer.
            Writer.WriteArray(Of Byte)(0L, System.Text.Encoding.ASCII.GetBytes(New String(Nothing, Me.MemoryBufferSize)), 0I, Me.MemoryBufferSize)

            ' Secondly write the byte-data. 
            Writer.WriteArray(Of Byte)(0L, Data, 0I, Data.Length)

        End Using ' Writer

    End Sub

    ''' <summary>
    ''' Reads a byte-sequence from a <see cref="MemoryMappedFile"/>.
    ''' </summary>
    ''' <param name="Name">Indicates an existing <see cref="MemoryMappedFile"/> assigned name.</param>
    ''' <param name="BufferLength">The buffer-length to read in.</param>
    ''' <returns>System.Byte().</returns>
    Private Function ReadMemoryMappedFile(ByVal Name As String, ByVal BufferLength As Integer) As Byte()

        Try
            Using MemoryFile As MemoryMappedFile =
                MemoryMappedFile.OpenExisting(Name, MemoryMappedFileRights.Read)

                Using Reader As MemoryMappedViewAccessor =
                    MemoryFile.CreateViewAccessor(0L, BufferLength, MemoryMappedFileAccess.Read)

                    Dim ReadBytes As Byte() = New Byte(BufferLength - 1I) {}
                    Reader.ReadArray(Of Byte)(0L, ReadBytes, 0I, ReadBytes.Length)
                    Return ReadBytes

                End Using ' Reader

            End Using ' MemoryFile

        Catch ex As IO.FileNotFoundException
            Throw
            Return Nothing

        End Try

    End Function

    ''' <summary>
    ''' Handles the 'Click' event of the 'btMakeFile' control.
    ''' </summary>
    Private Sub btMakeFile_Click() Handles btMakeFile.Click

        ' Get the byte-data to create the memory-mapped file.
        Dim WriteData As Byte() = System.Text.Encoding.ASCII.GetBytes(Me.strMessage)

        ' Create the memory-mapped file.
        Me.MakeMemoryMappedFile(Name:=Me.MemoryName, BufferLength:=Me.MemoryBufferSize, Data:=WriteData)

    End Sub

    ''' <summary>
    ''' Handles the 'Click' event of the 'btReadFile' control.
    ''' </summary>
    Private Sub btReadFile_Click() Handles btReadFile.Click


        Dim ReadBytes As Byte()

        Try ' Read the byte-sequence from memory.
            ReadBytes = ReadMemoryMappedFile(Name:=Me.MemoryName, BufferLength:=Me.MemoryBufferSize)

        Catch ex As IO.FileNotFoundException
            Me.tbReceptor.Text = "Memory-mapped file does not exist."
            Exit Sub

        End Try

        ' Convert the bytes to String.
        Dim Message As String = System.Text.Encoding.ASCII.GetString(ReadBytes.ToArray)

        ' Remove null chars (leading zero-bytes)
        Message = Message.Trim({ControlChars.NullChar})

        ' Print the message.
        tbReceptor.Text = Message

    End Sub

End Class

#End Region











' *************************
' This is the Application 2
' *************************

#Region " Imports "

Imports System.IO.MemoryMappedFiles

#End Region

#Region " Application 2 "

''' <summary>
''' Class MemoryMappedFile_Form2.
''' This should be the Class used to compile our first application.
''' </summary>
Public Class MemoryMappedFile_Form2

    ' The controls to create on execution-time.
    Dim WithEvents btMakeFile As New Button ' Writes the memory.
    Dim WithEvents btReadFile As New Button ' Reads the memory.
    Dim tbMessage As New TextBox ' Determines the string to map into memory.
    Dim tbReceptor As New TextBox ' Print the memory read's result.
    Dim lbInfoButtons As New Label ' Informs the user with a usage hint for the buttons.
    Dim lbInfotbMessage As New Label ' Informs the user with a usage hint for 'tbMessage'.

    ''' <summary>
    ''' Indicates the name of our memory-file.
    ''' </summary>
    Private ReadOnly MemoryName As String = "My Memory-File Name"

    ''' <summary>
    ''' Indicates the memory buffersize to store the <see cref="MemoryName"/>, in bytes.
    ''' </summary>
    Private ReadOnly MemoryBufferSize As Integer = 1024I

    ''' <summary>
    ''' Indicates the string to map in memory.
    ''' </summary>
    Private ReadOnly Property strMessage As String
        Get
            Return tbMessage.Text
        End Get
    End Property

    ''' <summary>
    ''' Initializes a new instance of the <see cref="MemoryMappedFile_Form2"/> class.
    ''' </summary>
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Set the properties of the controls.
        With lbInfotbMessage
            .Location = New Point(20, 10)
            .Text = "Type in this TextBox the message to write in memory:"
            .AutoSize = True
            ' .Size = tbReceptor.Size
        End With
        With tbMessage
            .Text = "Hello world from application two!"
            .Location = New Point(20, 30)
            .Size = New Size(310, Me.tbMessage.Height)
        End With
        With btMakeFile
            .Text = "Write Memory"
            .Size = New Size(130, 45)
            .Location = New Point(20, 50)
        End With
        With btReadFile
            .Text = "Read Memory"
            .Size = New Size(130, 45)
            .Location = New Point(200, 50)
        End With
        With tbReceptor
            .Location = New Point(20, 130)
            .Size = New Size(310, 100)
            .Multiline = True
        End With
        With lbInfoButtons
            .Location = New Point(tbReceptor.Location.X, tbReceptor.Location.Y - 30)
            .Text = "Press '" & btMakeFile.Text & "' button to create the memory file, that memory can be read from both applications."
            .AutoSize = False
            .Size = tbReceptor.Size
        End With

        ' Set the Form properties.
        With Me
            .Text = "Application 2"
            .Size = New Size(365, 300)
            .FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle
            .MaximizeBox = False
            .StartPosition = FormStartPosition.CenterScreen
        End With

        ' Add the controls on the UI.
        Me.Controls.AddRange({lbInfotbMessage, tbMessage, btMakeFile, btReadFile, tbReceptor, lbInfoButtons})

    End Sub

    ''' <summary>
    ''' Writes a byte sequence into a <see cref="MemoryMappedFile"/>.
    ''' </summary>
    ''' <param name="Name">Indicates the name to assign the <see cref="MemoryMappedFile"/>.</param>
    ''' <param name="BufferLength">Indicates the <see cref="MemoryMappedFile"/> buffer-length to write in.</param>
    ''' <param name="Data">Indicates the byte-data to write inside the <see cref="MemoryMappedFile"/>.</param>
    Private Sub MakeMemoryMappedFile(ByVal Name As String, ByVal BufferLength As Integer, ByVal Data As Byte())

        ' Create or open the memory-mapped file.
        Dim MessageFile As MemoryMappedFile =
            MemoryMappedFile.CreateOrOpen(Name, Me.MemoryBufferSize, MemoryMappedFileAccess.ReadWrite)

        ' Write the byte-sequence into memory.
        Using Writer As MemoryMappedViewAccessor =
            MessageFile.CreateViewAccessor(0L, Me.MemoryBufferSize, MemoryMappedFileAccess.ReadWrite)

            ' Firstly fill with null all the buffer.
            Writer.WriteArray(Of Byte)(0L, System.Text.Encoding.ASCII.GetBytes(New String(Nothing, Me.MemoryBufferSize)), 0I, Me.MemoryBufferSize)

            ' Secondly write the byte-data. 
            Writer.WriteArray(Of Byte)(0L, Data, 0I, Data.Length)

        End Using ' Writer

    End Sub

    ''' <summary>
    ''' Reads a byte-sequence from a <see cref="MemoryMappedFile"/>.
    ''' </summary>
    ''' <param name="Name">Indicates an existing <see cref="MemoryMappedFile"/> assigned name.</param>
    ''' <param name="BufferLength">The buffer-length to read in.</param>
    ''' <returns>System.Byte().</returns>
    Private Function ReadMemoryMappedFile(ByVal Name As String, ByVal BufferLength As Integer) As Byte()

        Try
            Using MemoryFile As MemoryMappedFile =
                MemoryMappedFile.OpenExisting(Name, MemoryMappedFileRights.Read)

                Using Reader As MemoryMappedViewAccessor =
                    MemoryFile.CreateViewAccessor(0L, BufferLength, MemoryMappedFileAccess.Read)

                    Dim ReadBytes As Byte() = New Byte(BufferLength - 1I) {}
                    Reader.ReadArray(Of Byte)(0L, ReadBytes, 0I, ReadBytes.Length)
                    Return ReadBytes

                End Using ' Reader

            End Using ' MemoryFile

        Catch ex As IO.FileNotFoundException
            Throw
            Return Nothing

        End Try

    End Function

    ''' <summary>
    ''' Handles the 'Click' event of the 'btMakeFile' control.
    ''' </summary>
    Private Sub btMakeFile_Click() Handles btMakeFile.Click

        ' Get the byte-data to create the memory-mapped file.
        Dim WriteData As Byte() = System.Text.Encoding.ASCII.GetBytes(Me.strMessage)

        ' Create the memory-mapped file.
        Me.MakeMemoryMappedFile(Name:=Me.MemoryName, BufferLength:=Me.MemoryBufferSize, Data:=WriteData)

    End Sub

    ''' <summary>
    ''' Handles the 'Click' event of the 'btReadFile' control.
    ''' </summary>
    Private Sub btReadFile_Click() Handles btReadFile.Click


        Dim ReadBytes As Byte()

        Try ' Read the byte-sequence from memory.
            ReadBytes = ReadMemoryMappedFile(Name:=Me.MemoryName, BufferLength:=Me.MemoryBufferSize)

        Catch ex As IO.FileNotFoundException
            Me.tbReceptor.Text = "Memory-mapped file does not exist."
            Exit Sub

        End Try

        ' Convert the bytes to String.
        Dim Message As String = System.Text.Encoding.ASCII.GetString(ReadBytes.ToArray)

        ' Remove null chars (leading zero-bytes)
        Message = Message.Trim({ControlChars.NullChar})

        ' Print the message.
        tbReceptor.Text = Message

    End Sub

End Class

#End Region
