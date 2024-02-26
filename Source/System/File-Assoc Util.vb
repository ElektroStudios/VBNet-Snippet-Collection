' ***********************************************************************
' Author   : Elektro
' Modified : 25-October-2015
' ***********************************************************************
' <copyright file="File-Assoc Util.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " Public Members Summary "

#Region " Types "

' FileAssocUtil.FileExtensionInfo <Serializable>

#End Region

#Region " Enumerations "

' FileAssocUtil.RegistryScope As Integer

#End Region

#Region " Properties "

' FileassocUtil.FileExtensionInfo.Command As String
' FileassocUtil.FileExtensionInfo.ContentType As String
' FileassocUtil.FileExtensionInfo.DdeApplication As String
' FileassocUtil.FileExtensionInfo.DdeCommand As String
' FileassocUtil.FileExtensionInfo.DdeIfExec As String
' FileassocUtil.FileExtensionInfo.DdeTopic As String
' FileassocUtil.FileExtensionInfo.DefaultIcon As String
' FileassocUtil.FileExtensionInfo.DelegateExecute As String
' FileassocUtil.FileExtensionInfo.DropTarget As String
' FileassocUtil.FileExtensionInfo.Executable As String
' FileassocUtil.FileExtensionInfo.FriendlyAppName As String
' FileassocUtil.FileExtensionInfo.FriendlyDocName As String
' FileassocUtil.FileExtensionInfo.InfoTip As String
' FileassocUtil.FileExtensionInfo.Name As String
' FileassocUtil.FileExtensionInfo.NoOpen As String
' FileassocUtil.FileExtensionInfo.QuickTip As String
' FileassocUtil.FileExtensionInfo.ShellExtension As String
' FileassocUtil.FileExtensionInfo.ShellNewValue As String
' FileassocUtil.FileExtensionInfo.SupportedUriProtocols As String
' FileassocUtil.FileExtensionInfo.TileInfo As String
' FileassocUtil.FileExtensionInfo.Max As String <Hidden>

#End Region

#Region " Functions "

' FileAssocUtil.GetFileExtensionInfo(String) As FileAssocUtil.FileExtensionInfo
' FileAssocUtil.IsRegistered(String) As Boolean

#End Region

#Region " Methods "

' FileAssocUtil.Register(FileAssocUtil.RegistryScope, String, String, Opt: String, Opt: String, Opt: Integer, Opt: String, Opt: String)

#End Region

#End Region

#Region " Usage Examples "

#Region " Register "

'FileAssocUtil.Register(scope:=FileAssocUtil.RegistryScope.CurrentUser,
'                       extensionName:=".elek",
'                       keyReferenceName:="ElektroFile",
'                       friendlyName:="Elektro File",
'                       defaultIcon:="%WinDir%\System32\Shell32.ico",
'                       iconIndex:=0,
'                       executable:="%WinDir%\notepad.exe",
'                       arguments:="""%1""")
'
'Dim isRegistered As Boolean = FileAssocUtil.IsRegistered(".elek")

#End Region

#Region " FileExtensionInfo "

'Dim feInfo As FileAssocUtil.FileExtensionInfo = FileAssocUtil.GetFileExtensionInfo(".wav")
'
'Dim sb As New StringBuilder
'With sb
'    .AppendLine(String.Format("Extension Name: {0}", feInfo.Name))
'    .AppendLine(String.Format("Friendly Doc Name: {0}", feInfo.FriendlyDocName))
'    .AppendLine(String.Format("Content Type: {0}", feInfo.ContentType))
'    .AppendLine(String.Format("Default Icon: {0}", feInfo.DefaultIcon))
'    .AppendLine("-----------------------------------------------------------")
'    .AppendLine(String.Format("Friendly App Name: {0}", feInfo.FriendlyAppName))
'    .AppendLine(String.Format("Executable: {0}", feInfo.Executable))
'    .AppendLine(String.Format("Command: {0}", feInfo.Command))
'    .AppendLine("-----------------------------------------------------------")
'    .AppendLine(String.Format("Drop Target: {0}", feInfo.DropTarget))
'    .AppendLine(String.Format("Info Tip: {0}", feInfo.InfoTip))
'    .AppendLine(String.Format("No Open: {0}", feInfo.NoOpen))
'    .AppendLine(String.Format("Shell Extension: {0}", feInfo.ShellExtension))
'    .AppendLine(String.Format("Shell New Value: {0}", feInfo.ShellNewValue))
'    .AppendLine("-----------------------------------------------------------")
'    .AppendLine(String.Format("Supported URI Protocols: {0}", feInfo.SupportedUriProtocols))
'    .AppendLine(String.Format("DDE Application: {0}", feInfo.DdeApplication))
'    .AppendLine(String.Format("DDE Command: {0}", feInfo.DdeCommand))
'    .AppendLine(String.Format("DDE If Exec: {0}", feInfo.DdeIfExec))
'    .AppendLine(String.Format("DDE Topic: {0}", feInfo.DdeTopic))
'End With
'
'MessageBox.Show(sb.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Information)

#End Region

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports Microsoft.Win32
Imports System
Imports System.ComponentModel
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Security.Permissions
Imports System.Text

#End Region

#Region " File-Assoc Util "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Contains related Windows file association utilities.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
<RegistryPermission(SecurityAction.Demand, Unrestricted:=True)>
Public Module FileAssocUtil

#Region " P/Invoking "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Platform Invocation methods (P/Invoke), access unmanaged code.
    ''' This class does not suppress stack walks for unmanaged code permission.
    ''' <see cref="System.Security.SuppressUnmanagedCodeSecurityAttribute"/> must not be applied to this class.
    ''' This class is for methods that can be used anywhere because a stack walk will be performed.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <remarks>
    ''' <see href="http://msdn.microsoft.com/en-us/library/ms182161.aspx"/>
    ''' </remarks>
    ''' ----------------------------------------------------------------------------------------------------
    Private NotInheritable Class NativeMethods

#Region " Functions "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Searches for and retrieves a file or protocol association-related string from the registry.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="flags">
        ''' The flags that can be used to control the search. It can be any combination of ASSOCF values, except that only one ASSOCF_INIT value can be included.
        ''' </param>
        ''' 
        ''' <param name="str">
        ''' An <see cref="ASSOCSTR"/> value that specifies the type of string that is to be returned.
        ''' </param>
        ''' 
        ''' <param name="pszAssoc">
        ''' A pointer to a null-terminated string that is used to determine the root key.
        ''' The following four types of strings can be used:
        ''' 
        ''' - A file name extension, such as '.txt'.
        ''' - A CLSID GUID in the standard '{GUID}' format.
        ''' - An application's ProgID, such as 'Word.Document.8.'
        ''' - The name of an application's .exe file (The ASSOCF_OPEN_BYEXENAME flag must be set in flags).
        ''' </param>
        ''' 
        ''' <param name="pszExtra">
        ''' An optional null-terminated string with additional information about the location of the string. 
        ''' It is typically set to a Shell verb such as 'open'. Set this parameter to <see langword="Nothing"/> if it is not used.
        ''' </param>
        ''' 
        ''' <param name="pszOut">
        ''' Pointer to a null-terminated string that, when this function returns successfully, receives the requested string. 
        ''' Set this parameter to <see langword="Nothing"/> to retrieve the required buffer size.
        ''' </param>
        ''' 
        ''' <param name="pcchOut">
        ''' A pointer to a value that, when calling the function, is set to the number of characters in the <paramref name="pszOut"/> buffer. 
        ''' When the function returns successfully, the value is set to the number of characters actually placed in the buffer.
        ''' 
        ''' If the ASSOCF_NOTRUNCATE flag is set in <paramref name="flags"/> and the 
        ''' buffer specified in <paramref name="pszOut"/> is too small, 
        ''' the function returns E_POINTER and the value is set to the required size of the buffer.
        ''' 
        ''' If <paramref name="pszOut"/> is <see langword="Nothing"/>,
        ''' the function returns S_FALSE and <paramref name="pcchOut"/> points to the required size, in characters, of the buffer.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <returns>
        ''' Returns a standard COM error value, including the following:
        ''' S_OK: Success.
        ''' E_POINTER: The pszOut buffer is too small to hold the entire string.
        ''' S_FALSE: <paramref name="pszOut"/> is <see langword="Nothing"/>, <paramref name="pcchOut"/> contains the required buffer size.
        ''' </returns>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb773471%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("Shlwapi.dll", SetLastError:=True, CharSet:=CharSet.Ansi, bestFitMapping:=False, throwOnUnmappableChar:=True)>
        Friend Shared Function AssocQueryString(
                               ByVal flags As AssocF,
                               ByVal str As AssocStr,
                               ByVal pszAssoc As String,
                               ByVal pszExtra As String,
                         <Out> ByVal pszOut As StringBuilder,
                  <[In]> <Out> ByRef pcchOut As UInteger
        ) As UInteger
        End Function

#End Region

#Region " Methods "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Notifies the system of an event that an application has performed. 
        ''' An application should use this function if it performs an action that may affect the Shell.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="wEventId">
        ''' Describes the event that has occurred. 
        ''' Typically, only one event is specified at a time. 
        ''' If more than one event is specified, the values contained in the 
        ''' <paramref name="dwItem1"/> and <paramref name="dwItem2"/> parameters must be the same, respectively, for all specified events.
        ''' </param>
        ''' 
        ''' <param name="uFlags">
        ''' Flags that, when combined bitwise with SHCNF_TYPE, 
        ''' indicate the meaning of the <paramref name="dwItem1"/> and <paramref name="dwItem2"/> parameters.
        ''' </param>
        ''' 
        ''' <param name="dwItem1">
        ''' Optional. 
        ''' First event-dependent value.
        ''' </param>
        ''' 
        ''' <param name="dwItem2">
        ''' Optional.
        ''' Second event-dependent value.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb762118%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <DllImport("shell32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
        Friend Shared Sub SHChangeNotify(
                      ByVal wEventId As FileAssocUtil.NativeMethods.SHChangeNotifyEventID,
                      ByVal uFlags As FileAssocUtil.NativeMethods.SHChangeNotifyFlags,
                      ByVal dwItem1 As IntPtr,
                      ByVal dwItem2 As IntPtr)
        End Sub

#End Region

#Region " Enumerations "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Provides information to the IQueryAssociations interface methods.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb762471%28v=vs.85%29.asp"/>x
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <Flags>
        Friend Enum AssocF As Integer

            ''' <summary>
            ''' None of the following options are set.
            ''' </summary>
            None = &H0

            ''' <summary>
            ''' Instructs IQueryAssociations interface methods not to map CLSID values to ProgID values.
            ''' </summary>
            InitNoRemapClsid = &H1

            ''' <summary>
            ''' Identifies the value of the pwszAssoc parameter of IQueryAssociations::Init as an executable file name. 
            ''' If this flag is not set, the root key will be set to the ProgID associated with the .exe key instead of the executable file's ProgID.
            ''' </summary>
            InitByExeName = &H2

            ''' <summary>
            ''' Identical to <see cref="AssocF.InitByExeName"/>.
            ''' </summary>
            OpenByExeName = &H2

            ''' <summary>
            ''' Specifies that when an IQueryAssociations method does not find the requested value under the root key, 
            ''' it should attempt to retrieve the comparable value from the * subkey.
            ''' </summary>
            InitDefaultToStar = &H4

            ''' <summary>
            ''' Specifies that when a IQueryAssociations method does not find the requested value under the root key, 
            ''' it should attempt to retrieve the comparable value from the Folder subkey.
            ''' </summary>
            InitDefaultToFolder = &H8

            ''' <summary>
            ''' Specifies that only HKEY_CLASSES_ROOT should be searched, and that HKEY_CURRENT_USER should be ignored.
            ''' </summary>
            NoUserSettings = &H10

            ''' <summary>
            ''' Specifies that the return string should not be truncated. 
            ''' Instead, return an error value and the required size for the complete string.
            ''' </summary>
            NoTruncate = &H20

            ''' <summary>
            ''' Instructs IQueryAssociations methods to verify that data is accurate. 
            ''' This setting allows IQueryAssociations methods to read data from the user's hard disk for verification. 
            ''' For example, they can check the friendly name in the registry against the one stored in the .exe file. 
            ''' Setting this flag typically reduces the efficiency of the method.
            ''' </summary>
            Verify = &H40

            ''' <summary>
            ''' Instructs IQueryAssociations methods to ignore Rundll.exe and return information about its target. 
            ''' Typically IQueryAssociations methods return information about the first .exe or .dll in a command string. 
            ''' If a command uses Rundll.exe, setting this flag tells the method to ignore Rundll.exe and return information about its target.
            ''' </summary>
            RemapRunDll = &H80

            ''' <summary>
            ''' Instructs IQueryAssociations methods not to fix errors in the registry, 
            ''' such as the friendly name of a function not matching the one found in the .exe file.
            ''' </summary>
            NoFixUps = &H100

            ''' <summary>
            ''' Specifies that the BaseClass value should be ignored.
            ''' </summary>
            IgnoreBaseClass = &H200

            ''' <summary>
            ''' (Introduced in Windows 7)
            ''' Specifies that the "Unknown" ProgID should be ignored; instead, fail.
            ''' </summary>
            IgnoreUnknown = &H400

            ''' <summary>
            ''' (Introduced in Windows 8)
            ''' Specifies that the supplied ProgID should be mapped using the system defaults, rather than the current user defaults.
            ''' </summary>
            InitFixedProgId = &H800

            ''' <summary>
            ''' (Introduced in Windows 8)
            ''' Specifies that the value is a protocol, and should be mapped using the current user defaults.
            ''' </summary>
            IsProtocol = &H1000

            ''' <summary>
            ''' (Introduced in Windows 8.1)
            ''' Specifies that the ProgID corresponds with a file extension based association. 
            ''' Use this flag together with <see cref="AssocF.InitFixedProgId"/>.
            ''' </summary>
            InitForFile = &H2000

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Used by IQueryAssociations::GetString to define the type of string that is to be returned.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb762475%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        Friend Enum AssocStr As Integer

            ''' <summary>
            ''' A command string associated with a Shell verb.
            ''' </summary>
            Command = 1

            ''' <summary>
            ''' An executable from a Shell verb command string. 
            ''' For example, this string is found as the (Default) value for a subkey such as 
            ''' HKEY_CLASSES_ROOT\ApplicationName\shell\Open\command. 
            ''' 
            ''' If the command uses Rundll.exe, set the <see cref="AssocF.RemapRunDll"/> flag in the flags parameter of 
            ''' IQueryAssociations::GetString to retrieve the target executable.
            ''' </summary>
            Executable = 2

            ''' <summary>
            ''' The friendly name of a document type.
            ''' </summary>
            FriendlyDocName = 3

            ''' <summary>
            ''' The friendly name of an executable file.
            ''' </summary>
            FriendlyAppName = 4

            ''' <summary>
            ''' Ignore the information associated with the open subkey.
            ''' </summary>
            NoOpen = 5

            ''' <summary>
            ''' Look under the ShellNew subkey.
            ''' </summary>
            ShellNewValue = 6

            ''' <summary>
            ''' A template for DDE commands.
            ''' </summary>
            DdeCommand = 7

            ''' <summary>
            ''' The DDE command to use to create a process.
            ''' </summary>
            DdeIfExec = 8

            ''' <summary>
            ''' The application name in a DDE broadcast.
            ''' </summary>
            DdeApplication = 9

            ''' <summary>
            ''' The topic name in a DDE broadcast.
            ''' </summary>
            DdeTopic = 10

            ''' <summary>
            ''' Corresponds to the InfoTip registry value. 
            ''' Returns an info tip for an item, or list of properties in the form of an IPropertyDescriptionList from which to 
            ''' create an info tip, such as when hovering the cursor over a file name. 
            ''' 
            ''' The list of properties can be parsed with PSGetPropertyDescriptionListFromString.
            ''' </summary>
            InfoTip = 11

            ''' <summary>
            ''' Corresponds to the QuickTip registry value. Same as <see cref="AssocStr.InfoTip"/>, 
            ''' except that it always returns a list of property names in the form of an IPropertyDescriptionList. 
            ''' The difference between this value and <see cref="AssocStr.InfoTip"/> is that this returns properties that are 
            ''' safe for any scenario that causes slow property retrieval, such as offline or slow networks. 
            ''' Some of the properties returned from <see cref="AssocStr.InfoTip"/> might not be appropriate for 
            ''' slow property retrieval scenarios. 
            ''' 
            ''' The list of properties can be parsed with PSGetPropertyDescriptionListFromString.
            ''' </summary>
            QuickTip = 12

            ''' <summary>
            ''' Corresponds to the TileInfo registry value. 
            ''' Contains a list of properties to be displayed for a particular file type in a Windows Explorer window that is in tile view. 
            ''' This is the same as <see cref="AssocStr.InfoTip"/>, but, like <see cref="AssocStr.QuickTip"/>, 
            ''' it also returns a list of property names in the form of an IPropertyDescriptionList. 
            ''' 
            ''' The list of properties can be parsed with PSGetPropertyDescriptionListFromString
            ''' </summary>
            TileInfo = 13

            ''' <summary>
            ''' Describes a general type of MIME file association, such as image and bmp, 
            ''' so that applications can make general assumptions about a specific file type.
            ''' </summary>
            ContentType = 14

            ''' <summary>
            ''' Returns the path to the icon resources to use by default for this association.
            ''' Positive numbers indicate an index into the dll's resource table, while negative numbers indicate a resource ID. 
            ''' 
            ''' An example of the syntax for the resource is "c:\myfolder\myfile.dll,-1".
            ''' </summary>
            DefaultIcon = 15

            ''' <summary>
            ''' For an object that has a Shell extension associated with it, you can use this to retrieve the CLSID of that Shell extension object 
            ''' by passing a string representation of the IID of the interface you want to retrieve as the pwszExtra parameter of IQueryAssociations::GetString. 
            ''' 
            ''' For example, if you want to retrieve a handler that implements the IExtractImage interface, 
            ''' you would specify "{BB2E617C-0920-11d1-9A0B-00C04FC2D6C1}", which is the IID of IExtractImage.
            ''' </summary>
            ShellExtension = 16

            ''' <summary>
            ''' For a verb invoked through COM and the IDropTarget interface, you can use this flag to retrieve the IDropTarget object's CLSID.
            ''' This CLSID is registered in the DropTarget subkey. The verb is specified in the pwszExtra parameter in the call to IQueryAssociations::GetString.
            ''' </summary>
            DropTarget = 17

            ''' <summary>
            ''' For a verb invoked through COM and the IExecuteCommand interface, you can use this flag to retrieve the IExecuteCommand object's CLSID. 
            ''' This CLSID is registered in the verb's command subkey as the DelegateExecute entry. 
            ''' The verb is specified in the pwszExtra parameter in the call to IQueryAssociations::GetString.
            ''' </summary>
            DelegateExecute = 18

            ''' <summary>
            ''' Introduced in Windows 8.
            ''' There is no official documentation for this flag.
            ''' </summary>
            SupportedUriProtocols = 19

            ''' <summary>
            ''' The maximum defined <see cref="AssocStr"/> value, used for validation purposes.
            ''' </summary>
            Max = 20

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Flags for <paramref name="wEventId"/> parameter of <see cref="FileAssocUtil.NativeMethods.SHChangeNotify"/> method.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb762118%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <Flags>
        Friend Enum SHChangeNotifyEventID As UInteger

            ' *****************************************************************************
            '                            WARNING!, NEED TO KNOW...
            '
            '  THIS ENUMERATION IS PARTIALLY DEFINED JUST FOR THE PURPOSES OF THIS PROJECT
            ' *****************************************************************************

            ''' <summary>
            ''' All events have occurred.
            ''' </summary>
            AllEvents = &H7FFFFFFFUI

            ''' <summary>
            ''' A file type association has changed. 
            ''' <see cref="FileAssocUtil.NativeMethods.SHChangeNotifyFlags.IdList"/> must be specified in the <paramref name="uFlags"/> parameter.
            ''' 
            ''' <paramref name="dwItem1"/> and <paramref name="dwItem2"/> are not used and must be set as <see cref="IntPtr.Zero"/>.
            ''' </summary>
            AssocChanged = &H8000000UI

        End Enum

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Flags for <paramref name="uFlags"/> parameter of <see cref="FileAssocUtil.NativeMethods.SHChangeNotify"/> method.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <remarks>
        ''' <see href="http://msdn.microsoft.com/en-us/library/windows/desktop/bb762118%28v=vs.85%29.aspx"/>
        ''' </remarks>
        ''' ----------------------------------------------------------------------------------------------------
        <Flags>
        Friend Enum SHChangeNotifyFlags As UInteger

            ' *****************************************************************************
            '                            WARNING!, NEED TO KNOW...
            '
            '  THIS ENUMERATION IS PARTIALLY DEFINED JUST FOR THE PURPOSES OF THIS PROJECT
            ' *****************************************************************************

            ''' <summary>
            ''' <paramref name="dwItem1"/> and <paramref name="dwItem2"/> are the addresses of 'ITEMIDLIST' structures that
            ''' represent the item(s) affected by the change.
            ''' 
            ''' Each 'ITEMIDLIST' must be relative to the desktop folder.
            ''' </summary>
            IdList = &H0UI

        End Enum

#End Region

    End Class

#End Region

#Region " Types "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Defines the system information of a file extension.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <Serializable>
    Public NotInheritable Class FileExtensionInfo

#Region " Properties "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the file extension name.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The file extension name.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public ReadOnly Property Name As String
            <DebuggerStepThrough>
            Get
                Return Me.nameB
            End Get
        End Property
        ''' <summary>
        ''' ( Backing field )
        ''' The file extension name.
        ''' </summary>
        Private ReadOnly nameB As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a command string associated with a Shell verb.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' A command string associated with a Shell verb.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property Command As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets an executable from a Shell verb command string. 
        ''' For example, this string is found as the (Default) value for a subkey such as 
        ''' HKEY_CLASSES_ROOT\ApplicationName\shell\Open\command.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' An executable from a Shell verb command string.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property Executable As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the friendly name of a document type.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The friendly name of a document type.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property FriendlyDocName As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the friendly name of an executable file.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The friendly name of an executable type.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property FriendlyAppName As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a value indicating whether the association ignores the information associated with the 'open' subkey command.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' A value indicating whether the association ignores the information associated with the 'open' subkey command.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property NoOpen As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Look under the ShellNew subkey.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property ShellNewValue As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a template for DDE commands.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' A template for DDE commands.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property DdeCommand As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the DDE command to use to create a process.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The DDE command to use to create a process.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property DdeIfExec As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the application name in a DDE broadcast.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The application name in a DDE broadcast.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property DdeApplication As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the topic name in a DDE broadcast.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The topic name in a DDE broadcast.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property DdeTopic As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Corresponds to the InfoTip registry value. 
        ''' Gets an info tip for an item, or list of properties in the form of an IPropertyDescriptionList from which to create an info tip, 
        ''' such as when hovering the cursor over a file name. 
        ''' 
        ''' The list of properties can be parsed with PSGetPropertyDescriptionListFromString.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property InfoTip As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Corresponds to the QuickTip registry value. Same as <see cref="FileAssocUtil.FileExtensionInfo.InfoTip"/>, 
        ''' except that it always returns a list of property names in the form of an IPropertyDescriptionList. 
        ''' 
        ''' The difference between this value and <see cref="FileAssocUtil.FileExtensionInfo.InfoTip"/> is 
        ''' that this returns properties that are safe for any scenario that causes slow property retrieval, 
        ''' such as offline or slow networks. 
        ''' 
        ''' Some of the properties returned from <see cref="FileAssocUtil.FileExtensionInfo.InfoTip"/> might not be appropriate for 
        ''' slow property retrieval scenarios. 
        ''' 
        ''' The list of properties can be parsed with PSGetPropertyDescriptionListFromString.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property QuickTip As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Corresponds to the TileInfo registry value. 
        ''' Contains a list of properties to be displayed for a particular file type in a Windows Explorer window that is in tile view. 
        ''' 
        ''' This is the same as <see cref="FileAssocUtil.FileExtensionInfo.InfoTip"/>, but, 
        ''' like <see cref="FileAssocUtil.FileExtensionInfo.QuickTip"/>, it also returns a list of property names in the 
        ''' form of an IPropertyDescriptionList. 
        ''' 
        ''' The list of properties can be parsed with PSGetPropertyDescriptionListFromString
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property TileInfo As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets a general type of MIME file association, such as 'image' and 'bmp', 
        ''' so that applications can make general assumptions about a specific file type.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' A general type of MIME file association, such as 'image' and 'bmp'.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property ContentType As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Gets the path to the icon resources to use by default for this association.
        ''' Positive numbers indicate an index into the dll's resource table, while negative numbers indicate a resource ID. 
        ''' 
        ''' An example of the syntax for the resource is "c:\myfolder\myfile.dll,-1".
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' A command string associated with a Shell verb.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property DefaultIcon As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' For an object that has a Shell extension associated with it, you can use this to retrieve the CLSID of 
        ''' that Shell extension object.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The CLSID of the associated Shell extension object.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property ShellExtension As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' For a verb invoked through COM and the IDropTarget interface, you can use this flag to retrieve the IDropTarget object's CLSID.
        ''' This CLSID is registered in the DropTarget subkey. The verb is specified in the pwszExtra parameter in the 
        ''' call to IQueryAssociations::GetString.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The CLSID of the associated IDropTarget object.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property DropTarget As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' For a verb invoked through COM and the IExecuteCommand interface, you can use this flag to retrieve the 
        ''' IExecuteCommand object's CLSID. 
        ''' 
        ''' This CLSID is registered in the verb's command subkey as the DelegateExecute entry. 
        ''' The verb is specified in the pwszExtra parameter in the call to IQueryAssociations::GetString.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <value>
        ''' The CLSID of the associated IExecuteCommand object.
        ''' </value>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property DelegateExecute As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' (Introduced in Windows 8)
        ''' There is no official documentation for this member.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Public Property SupportedUriProtocols As String

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' The maximum defined <see cref="FileAssocUtil.NativeMethods.AssocStr"/> value, used for validation purposes.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        <EditorBrowsable(EditorBrowsableState.Never)>
        Public Property Max As String

#End Region

#Region " Constructors "

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Prevents a default instance of the <see cref="FileExtensionInfo"/> class from being created.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        Private Sub New()
        End Sub

        ''' ----------------------------------------------------------------------------------------------------
        ''' <summary>
        ''' Initializes a new instance of the <see cref="FileExtensionInfo"/> class.
        ''' </summary>
        ''' ----------------------------------------------------------------------------------------------------
        ''' <param name="extensionName">
        ''' The name of the file extension.
        ''' </param>
        ''' ----------------------------------------------------------------------------------------------------
        Public Sub New(ByVal extensionName As String)

            Me.nameB = extensionName.TrimStart({"."c}).ToLower ' Remove extension dot and fix letter casing.

        End Sub

#End Region

    End Class

#End Region

#Region " Enumerations "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Specified a registry scope (root key).
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    Public Enum RegistryScope As Integer

        ''' <summary>
        ''' This reffers to the commonly known "HKLM" or "HKEY_LOCAL_MACHINE" registry root key.
        ''' Changes made on this registry root key will affect all users.
        ''' </summary>
        Machine = 0

        ''' <summary>
        ''' Current User, this reffers to the commonly known "HKCU" or "HKEY_CURRENT_USER" registry root key.
        ''' Changes made on this registry root key will affect the current user.
        ''' </summary>
        CurrentUser = 1

    End Enum

#End Region

#Region " Public Methods "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the system information of the specified file extension.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim feInfo As FileAssocUtil.FileExtensionInfo = FileAssocUtil.GetFileExtensionInfo(".wav")
    ''' 
    ''' Dim sb As New StringBuilder
    ''' With sb
    '''     .AppendLine(String.Format("Extension Name: {0}", feInfo.Name))
    '''     .AppendLine(String.Format("Friendly Doc Name: {0}", feInfo.FriendlyDocName))
    '''     .AppendLine(String.Format("Content Type: {0}", feInfo.ContentType))
    '''     .AppendLine(String.Format("Default Icon: {0}", feInfo.DefaultIcon))
    '''     .AppendLine("-----------------------------------------------------------")
    '''     .AppendLine(String.Format("Friendly App Name: {0}", feInfo.FriendlyAppName))
    '''     .AppendLine(String.Format("Executable: {0}", feInfo.Executable))
    '''     .AppendLine(String.Format("Command: {0}", feInfo.Command))
    '''     .AppendLine("-----------------------------------------------------------")
    '''     .AppendLine(String.Format("Drop Target: {0}", feInfo.DropTarget))
    '''     .AppendLine(String.Format("Info Tip: {0}", feInfo.InfoTip))
    '''     .AppendLine(String.Format("No Open: {0}", feInfo.NoOpen))
    '''     .AppendLine(String.Format("Shell Extension: {0}", feInfo.ShellExtension))
    '''     .AppendLine(String.Format("Shell New Value: {0}", feInfo.ShellNewValue))
    '''     .AppendLine("-----------------------------------------------------------")
    '''     .AppendLine(String.Format("Supported URI Protocols: {0}", feInfo.SupportedUriProtocols))
    '''     .AppendLine(String.Format("DDE Application: {0}", feInfo.DdeApplication))
    '''     .AppendLine(String.Format("DDE Command: {0}", feInfo.DdeCommand))
    '''     .AppendLine(String.Format("DDE If Exec: {0}", feInfo.DdeIfExec))
    '''     .AppendLine(String.Format("DDE Topic: {0}", feInfo.DdeTopic))
    ''' End With
    ''' 
    ''' MessageBox.Show(sb.ToString, "", MessageBoxButtons.OK, MessageBoxIcon.Information)
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="extensionName">
    ''' The extension name (eg: .txt).
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <exception cref="ArgumentNullException">
    ''' extensionName
    ''' </exception>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' A <see cref="FileAssocUtil.FileExtensionInfo"/> object that contains the file extension info.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Function GetFileExtensionInfo(ByVal extensionName As String) As FileExtensionInfo

        If String.IsNullOrWhiteSpace(extensionName) Then
            Throw New ArgumentNullException(paramName:="extensionName")

        Else ' Fix extension dot.
            extensionName = extensionName.TrimStart({"."c}).Insert(0, "."c)

        End If

        Dim feInfo As New FileExtensionInfo(extensionName)
        Dim sb As New StringBuilder
        Dim pcchout As UInteger

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.Command, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.Command, extensionName, Nothing, sb, pcchout)
        feInfo.Command = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.ContentType, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.ContentType, extensionName, Nothing, sb, pcchout)
        feInfo.ContentType = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DdeApplication, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DdeApplication, extensionName, Nothing, sb, pcchout)
        feInfo.DdeApplication = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DdeCommand, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DdeCommand, extensionName, Nothing, sb, pcchout)
        feInfo.DdeCommand = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DdeIfExec, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DdeIfExec, extensionName, Nothing, sb, pcchout)
        feInfo.DdeIfExec = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DdeTopic, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DdeTopic, extensionName, Nothing, sb, pcchout)
        feInfo.DdeTopic = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DefaultIcon, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DefaultIcon, extensionName, Nothing, sb, pcchout)
        feInfo.DefaultIcon = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DelegateExecute, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DelegateExecute, extensionName, Nothing, sb, pcchout)
        feInfo.DelegateExecute = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DropTarget, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DropTarget, extensionName, Nothing, sb, pcchout)
        feInfo.DropTarget = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.Executable, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.Executable, extensionName, Nothing, sb, pcchout)
        feInfo.Executable = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.FriendlyAppName, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.FriendlyAppName, extensionName, Nothing, sb, pcchout)
        feInfo.FriendlyAppName = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.FriendlyDocName, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.FriendlyDocName, extensionName, Nothing, sb, pcchout)
        feInfo.FriendlyDocName = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.InfoTip, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.InfoTip, extensionName, Nothing, sb, pcchout)
        feInfo.InfoTip = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.Max, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.Max, extensionName, Nothing, sb, pcchout)
        feInfo.Max = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.NoOpen, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.NoOpen, extensionName, Nothing, sb, pcchout)
        feInfo.NoOpen = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.QuickTip, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.QuickTip, extensionName, Nothing, sb, pcchout)
        feInfo.QuickTip = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.ShellExtension, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.ShellExtension, extensionName, Nothing, sb, pcchout)
        feInfo.ShellExtension = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.ShellNewValue, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.ShellNewValue, extensionName, Nothing, sb, pcchout)
        feInfo.ShellNewValue = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.SupportedUriProtocols, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.SupportedUriProtocols, extensionName, Nothing, sb, pcchout)
        feInfo.SupportedUriProtocols = sb.ToString()
        sb.Clear()

        pcchout = 0UI
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.TileInfo, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.TileInfo, extensionName, Nothing, sb, pcchout)
        feInfo.TileInfo = sb.ToString()
        sb.Clear()

        Return feInfo

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Determines whether the specified file extension is registered in the current system.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' Dim isRegistered As Boolean = FileAssocUtil.IsRegistered(".ext")
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="extensionName">
    ''' The extension name (eg: .txt).
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <exception cref="ArgumentNullException">
    ''' extensionName
    ''' </exception>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' <see langword="True"/> if the file extension is registered, otherwise, <see langword="False"/>.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Function IsRegistered(ByVal extensionName As String) As Boolean

        If String.IsNullOrWhiteSpace(extensionName) Then
            Throw New ArgumentNullException(paramName:="extensionName")

        Else ' Fix extension dot.
            extensionName = extensionName.TrimStart({"."c}).Insert(0, "."c)

        End If

        Dim sb As New StringBuilder
        Dim pcchout As UInteger = 0UI

        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DdeApplication, extensionName, Nothing, Nothing, pcchout)
        sb.Capacity = CInt(pcchout)
        NativeMethods.AssocQueryString(NativeMethods.AssocF.Verify, NativeMethods.AssocStr.DdeApplication, extensionName, Nothing, sb, pcchout)
        Return Not sb.ToString().Equals("OpenWith", StringComparison.OrdinalIgnoreCase)

    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Registers or modifies a file extension in the current system.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' FileAssocUtil.Register(scope:=FileAssocUtil.RegistryScope.CurrentUser,
    '''                extensionName:=".elek",
    '''                keyReferenceName:="ElektroFile",
    '''                friendlyName:="Elektro File",
    '''                defaultIcon:="%WinDir%\System32\Shell32.ico",
    '''                iconIndex:=0,
    '''                executable:="%WinDir%\notepad.exe",
    '''                arguments:="""%1""")
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="scope">
    ''' The registry scope that will owns the file association.
    ''' </param>
    ''' 
    ''' <param name="extensionName">
    ''' The extension name (eg: .txt).
    ''' </param>
    ''' 
    ''' <param name="keyReferenceName">
    ''' The name of the registry key to reference.
    ''' </param>
    ''' 
    ''' <param name="friendlyName">
    ''' The friendly name for this file extension (visible in Windows Esplorer).
    ''' </param>
    ''' 
    ''' <param name="defaultIcon">
    ''' The icon filepath.
    ''' </param>
    ''' 
    ''' <param name="iconIndex">
    ''' The index image of the icon resource.
    ''' </param>
    ''' 
    ''' <param name="executable">
    ''' The executable path that will open the file.
    ''' </param>
    ''' 
    ''' <param name="arguments">
    ''' The executable arguments,
    ''' normally this value is set as '"%1"' to take the filepath as the only one argument, 
    ''' or '"%1" %*' to take the filepath as first argument and additional arguments.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <exception cref="ArgumentNullException">
    ''' extensionName
    ''' </exception>
    ''' 
    ''' <exception cref="ArgumentException">
    ''' Invalid enumeration value.;scope
    ''' </exception>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Sub Register(ByVal scope As FileAssocUtil.RegistryScope,
                               ByVal extensionName As String,
                               ByVal keyReferenceName As String,
                               Optional ByVal friendlyName As String = "",
                               Optional ByVal defaultIcon As String = "%WinDir%\System32\shell32.dll",
                               Optional ByVal iconIndex As Integer = 0,
                               Optional ByVal executable As String = "%WinDir%\System32\OpenWith.exe",
                               Optional ByVal arguments As String = """%1""")

        If String.IsNullOrWhiteSpace(extensionName) Then
            Throw New ArgumentNullException(paramName:="extensionName")

        Else ' Fix extension dot.
            extensionName = extensionName.TrimStart({"."c}).Insert(0, "."c)

        End If

        Dim regKey As RegistryKey

        Select Case scope

            Case FileAssocUtil.RegistryScope.Machine
                regKey = Registry.LocalMachine

            Case FileAssocUtil.RegistryScope.CurrentUser
                regKey = Registry.CurrentUser

            Case Else
                Throw New ArgumentException(message:="Invalid enumeration value.", paramName:="scope")

        End Select

        Using regKey

            ' Create sub-key 'HKxx\Software\Classes'
            regKey.CreateSubKey("Software\Classes")

            ' Create sub-key 'HKxx\Software\Classes\{.ext}'
            regKey.OpenSubKey("Software\Classes", writable:=True).
                   CreateSubKey(extensionName)

            ' Set value data 'HKxx\Software\Classes\{.ext}\@Default={keyReferenceName}'
            regKey.OpenSubKey(String.Format("Software\Classes\{0}", extensionName), writable:=True).
                   SetValue("", keyReferenceName, RegistryValueKind.String)

            ' Create sub-key 'HKxx\Software\Classes\{KeyRefrenceName}'.
            regKey.OpenSubKey("Software\Classes", writable:=True).
                   CreateSubKey(keyReferenceName)

            ' Set value data 'HKxx\Software\Classes\{KeyRefrenceName}\@Default={friendlyName}'
            regKey.OpenSubKey(String.Format("Software\Classes\{0}", keyReferenceName), writable:=True).
                   SetValue("", friendlyName, RegistryValueKind.String)

            ' Create sub-key 'HKxx\Software\Classes\{KeyRefrenceName}\DefaultIcon'.
            regKey.OpenSubKey(String.Format("Software\Classes\{0}", keyReferenceName), writable:=True).
                   CreateSubKey("DefaultIcon")

            ' Set value data 'HKxx\Software\Classes\{KeyRefrenceName}\DefaultIcon\@Default={defaultIcon},{iconIndex}'.
            regKey.OpenSubKey(String.Format("Software\Classes\{0}\DefaultIcon", keyReferenceName), writable:=True).
                   SetValue("", Environment.ExpandEnvironmentVariables(String.Format("""{0}"",{1}", defaultIcon, CStr(iconIndex))), RegistryValueKind.String)

            ' Create sub-key 'HKxx\Software\Classes\{KeyRefrenceName}\Shell\Open\Command'.
            regKey.OpenSubKey(String.Format("Software\Classes\{0}", keyReferenceName), writable:=True).
                   CreateSubKey("Shell\Open\Command")

            ' Set value data 'HKxx\Software\Classes\{KeyRefrenceName}\Shell\Open\Command\@Default={executable} {arguments}'.
            regKey.OpenSubKey(String.Format("Software\Classes\{0}\Shell\Open\Command", keyReferenceName), writable:=True).
                   SetValue("", Environment.ExpandEnvironmentVariables(String.Format("""{0}"" {1}", executable, arguments)), RegistryValueKind.String)

            ' Delete the system 'OpenWith' override.
            regKey.DeleteSubKeyTree(String.Format("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\{0}\OpenWithProgIds", extensionName), throwOnMissingSubKey:=False)
            regKey.DeleteSubKeyTree(String.Format("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\{0}\UserChoice", extensionName), throwOnMissingSubKey:=False)

            ' Set the ProgId.
            regKey.CreateSubKey(String.Format("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\{0}\UserChoice", extensionName)).
                   SetValue("ProgId", keyReferenceName, RegistryValueKind.String)

            regKey.CreateSubKey(String.Format("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\{0}\OpenWithProgids", extensionName)).
                   SetValue(keyReferenceName, New Byte() {}, RegistryValueKind.None)

            regKey.CreateSubKey(String.Format("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\{0}\OpenWithList", extensionName)).
                   SetValue("MRUList", "a", RegistryValueKind.String)

            regKey.CreateSubKey(String.Format("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\{0}\OpenWithList", extensionName)).
                   SetValue("a", Path.GetFileName(executable), RegistryValueKind.String)

            ' Notify the system that a file association has been added/changed to update the changes (on Windows Explorer).
            NativeMethods.SHChangeNotify(FileAssocUtil.NativeMethods.SHChangeNotifyEventID.AssocChanged,
                                         FileAssocUtil.NativeMethods.SHChangeNotifyFlags.IdList,
                                         dwItem1:=IntPtr.Zero, dwItem2:=IntPtr.Zero)

        End Using

    End Sub

#End Region

#Region " Hidden methods "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Determines whether the specified System.Object instances are considered equal.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <EditorBrowsable(EditorBrowsableState.Never)>
    Public Function Equals(ByVal obj As Object) As Boolean
        Return MyBase.Equals(obj)
    End Function

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Determines whether the specified System.Object instances are the same instance.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    <EditorBrowsable(EditorBrowsableState.Never)>
    Public Function ReferenceEquals(ByVal objA As Object, ByVal objB As Object) As Boolean
        Return Nothing
    End Function

#End Region

End Module

#End Region
