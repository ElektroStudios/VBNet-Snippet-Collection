
# System snippets category
These snippets are intended to help managing the (operating)system in some aspects.

# General Information about this repository
 - .snippet files contains the formatted code that can be loaded through Visual Studio's code editor context menu.
 - .vb files contains the raw code that can be copied then pasted in any project.
 - Each .snippet and .vb file contains a #Region section and/or Xml documentation with code examples.
 
Feel free to use and/or modify any file of this repository.

If you like the job I've done, then please contribute with improvements to these snippets or by adding new ones.

#Environment Util
Huge class that contains related Windows environment utilities.

The 'EnvironmentVariables' class manages the environment variables, it can find, register and unregister vars.

The 'OS' class is very useful to tweak some Windows values, it also notifies the system about changes to perform a friendly update.

The 'FileSystem' class has some validation methods for filenames, and also can invoke item verbs.

The 'Shell' class pins and unpins items in taskbar and startmenu, it also can show/hide taskbar and startmenu, can get all the Explorer windows instances, can refresh Explorer windows, or add a file to recent docs, and much more.

The 'Theming' class manages the current theme and wallpaper, it can set other theme or wall, or change the system cursors.

Public Members Summary

 - Child Classes
   - EnvironmentUtil.EnvironmentVariables
   - EnvironmentUtil.FileSystem
   - EnvironmentUtil.OS
   - EnvironmentUtil.Programs
   - EnvironmentUtil.Shell
   - EnvironmentUtil.Shell.Desktop
   - EnvironmentUtil.Shell.Explorer
   - EnvironmentUtil.Shell.StartMenu
   - EnvironmentUtil.Shell.TaskBar
   - EnvironmentUtil.Theming

 - Enumerations
   - EnvironmentUtil.EnvironmentScope
   - EnvironmentUtil.OS.Architecture
   - EnvironmentUtil.Theming.CursorType
   - EnvironmentUtil.Theming.WallpaperStyle

 - Properties
   - EnvironmentUtil.EnvironmentVariables.CurrentVariables(EnvironmentUtil.EnvironmentScope) As ReadOnlyCollection(Of EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo)
   - EnvironmentUtil.OS.ActiveWindowTrackingEnabled As Boolean
   - EnvironmentUtil.OS.ActiveWindowTrackingTimeout As UShort
   - EnvironmentUtil.OS.BeepEnabled As Boolean
   - EnvironmentUtil.OS.BlockSendInputResetsEnabled As Boolean
   - EnvironmentUtil.OS.BorderMultiplierFactor As Integer
   - EnvironmentUtil.OS.CaretWidth As Integer
   - EnvironmentUtil.OS.CleartypeEnabled As Boolean
   - EnvironmentUtil.OS.ClientAreaAnimationEnabled As Boolean
   - EnvironmentUtil.OS.ComboBoxAnimationEnabled As Boolean
   - EnvironmentUtil.OS.CurrentArchitecture() As EnvironmentUtil.OS.Architecture
   - EnvironmentUtil.OS.CursorShadowEnabled As Boolean
   - EnvironmentUtil.OS.DoubleClickSize As Size
   - EnvironmentUtil.OS.DoubleClickTime As Integer
   - EnvironmentUtil.OS.DragFullWindowsEnabled As Boolean
   - EnvironmentUtil.OS.DragSize As Size
   - EnvironmentUtil.OS.DropShadowEnabled As Boolean
   - EnvironmentUtil.OS.FlatMenuEnabled As Boolean
   - EnvironmentUtil.OS.FocusBorderSize As Size
   - EnvironmentUtil.OS.FontSmoothingContrast As Integer
   - EnvironmentUtil.OS.FontSmoothingEnabled As Boolean
   - EnvironmentUtil.OS.ForegroundFlashCount As UShort
   - EnvironmentUtil.OS.ForegroundLockTimeout As UShort
   - EnvironmentUtil.OS.HotTrackingEnabled As Boolean
   - EnvironmentUtil.OS.HungAppTimeout As Integer
   - EnvironmentUtil.OS.IconSpacing As Size
   - EnvironmentUtil.OS.IconTitleWrappingEnabled As Boolean
   - EnvironmentUtil.OS.KeyboardDelay As Integer
   - EnvironmentUtil.OS.KeyboardSpeed As Integer
   - EnvironmentUtil.OS.ListBoxSmoothScrollingEnabled As Boolean
   - EnvironmentUtil.OS.MenuAccessKeysUnderlined As Boolean
   - EnvironmentUtil.OS.MenuAnimationEnabled As Boolean
   - EnvironmentUtil.OS.MenuFadeEnabled As Boolean
   - EnvironmentUtil.OS.MenuShowDelay As Integer
   - EnvironmentUtil.OS.MessageDuration As Long
   - EnvironmentUtil.OS.MouseButtonsSwapEnabled As Boolean
   - EnvironmentUtil.OS.MouseClickLockEnabled As Boolean
   - EnvironmentUtil.OS.MouseClickLockTime As Integer
   - EnvironmentUtil.OS.MouseHoverSize As Size
   - EnvironmentUtil.OS.MouseHoverTime As Integer
   - EnvironmentUtil.OS.MouseSonarEnabled As Boolean
   - EnvironmentUtil.OS.MouseSpeed As Integer
   - EnvironmentUtil.OS.MouseTrailAmount As Integer
   - EnvironmentUtil.OS.MouseVanishEnabled As Boolean
   - EnvironmentUtil.OS.MouseWheelScrollLines As Integer
   - EnvironmentUtil.OS.OverlappedContentEnabled As Boolean
   - EnvironmentUtil.OS.PopupMenuAlignment As LeftRightAlignment
   - EnvironmentUtil.OS.ScreensaverEnabled As Boolean
   - EnvironmentUtil.OS.ScreensaverPath As String
   - EnvironmentUtil.OS.ScreensaverTimeout As Integer
   - EnvironmentUtil.OS.ScreensaveSecureEnabled As Boolean
   - EnvironmentUtil.OS.SelectionFadeEnabled As Boolean
   - EnvironmentUtil.OS.SnapToDefaultEnabled As Boolean
   - EnvironmentUtil.OS.SystemDateTime As Date
   - EnvironmentUtil.OS.SystemLanguageBarEnabled As Boolean
   - EnvironmentUtil.OS.TitleBarGradientEnabled As Boolean
   - EnvironmentUtil.OS.ToolTipAnimationEnabled As Boolean
   - EnvironmentUtil.OS.UIEffectsEnabled As Boolean
   - EnvironmentUtil.OS.WaitToKillAppTimeout As Integer
   - EnvironmentUtil.OS.WaitToKillServiceTimeout As Integer
   - EnvironmentUtil.OS.WheelscrollChars As Integer
   - EnvironmentUtil.Programs.DefaultWebBrowser() As String
   - EnvironmentUtil.Programs.IExplorerVersion() As Version
   - EnvironmentUtil.Shell.Explorer.ExplorerWindows As ReadOnlyCollection(Of ShellBrowserWindow)
   - EnvironmentUtil.Shell.Explorer.ExplorerWindowsFolders As ReadOnlyCollection(Of Shell32.Folder2)
   - EnvironmentUtil.Shell.TaskBar.ClassName() As String
   - EnvironmentUtil.Shell.TaskBar.Hwnd() As Intptr
   - EnvironmentUtil.Theming.AeroEnabled() As Boolean
   - EnvironmentUtil.Theming.AeroSupported() As Boolean
   - EnvironmentUtil.Theming.CurrentTheme() As EnvironmentUtil.Theming.ThemeInfo
   - EnvironmentUtil.Theming.CurrentWallpaper() As String
   - EnvironmentUtil.Theming.WallpaperAsJpegIsSupported() As Boolean
   - EnvironmentUtil.Theming.WallpaperStylesFitFillAreSupported() As Boolean

 - Types
   - EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo
   - EnvironmentUtil.Theming.ThemeInfo

 - Functions
   - EnvironmentUtil.EnvironmentVariables.GetValue(EnvironmentUtil.EnvironmentScope, String, Boolean) As String
   - EnvironmentUtil.EnvironmentVariables.GetVariableInfo(EnvironmentUtil.EnvironmentScope, String, Boolean) As EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo
   - EnvironmentUtil.FileSystem.GetItemVerbs(String) As IEnumerable(Of FolderItemVerb)
   - EnvironmentUtil.FileSystem.ItemNameIsInvalid(String) As Boolean
   - EnvironmentUtil.FileSystem.ItemNameOrPathIsInvalid(String) As Boolean
   - EnvironmentUtil.FileSystem.ItemPathIsInvalid(String) As Boolean

 - Methods
   - EnvironmentUtil.EnvironmentVariables.RegisterVariable(EnvironmentUtil.EnvironmentScope, EnvironmentUtil.EnvironmentVariables.EnvironmentVariableInfo, Boolean)
   - EnvironmentUtil.EnvironmentVariables.RegisterVariable(EnvironmentUtil.EnvironmentScope, String, String, Boolean)
   - EnvironmentUtil.EnvironmentVariables.UnregisterVariable(EnvironmentUtil.EnvironmentScope, String, Boolean)
   - EnvironmentUtil.FileSystem.InvokeItemVerb(String, String)
   - EnvironmentUtil.OS.NotifyDirectoryAttributesChanged(String)
   - EnvironmentUtil.OS.NotifyDirectoryCreated(String)
   - EnvironmentUtil.OS.NotifyDirectoryDeleted(String)
   - EnvironmentUtil.OS.NotifyDirectoryRenamed(String, String)
   - EnvironmentUtil.OS.NotifyDriveAdded(String, Boolean)
   - EnvironmentUtil.OS.NotifyDriveRemoved(String)
   - EnvironmentUtil.OS.NotifyFileAssociationChanged()
   - EnvironmentUtil.OS.NotifyFileAttributesChanged(String)
   - EnvironmentUtil.OS.NotifyFileCreated(String)
   - EnvironmentUtil.OS.NotifyFileDeleted(String)
   - EnvironmentUtil.OS.NotifyFileRenamed(String, String)
   - EnvironmentUtil.OS.NotifyFreespaceChanged(String)
   - EnvironmentUtil.OS.NotifyMediaInserted(String)
   - EnvironmentUtil.OS.NotifyMediaRemoved(String)
   - EnvironmentUtil.OS.NotifyNetworkFolderShared(String)
   - EnvironmentUtil.OS.NotifyNetworkFolderUnshared(String)
   - EnvironmentUtil.OS.NotifyUpdateDirectory(String)
   - EnvironmentUtil.OS.NotifyUpdateImage()
   - EnvironmentUtil.OS.ReloadSystemCursors()
   - EnvironmentUtil.OS.ReloadSystemIcons()
   - EnvironmentUtil.OS.RunDateTime()
   - EnvironmentUtil.OS.RunExecuteDialog()
   - EnvironmentUtil.OS.RunFindComputer()
   - EnvironmentUtil.OS.RunFindFiles()
   - EnvironmentUtil.OS.RunFindPrinter()
   - EnvironmentUtil.OS.RunHelpCenter()
   - EnvironmentUtil.OS.RunSearchCommand()
   - EnvironmentUtil.OS.RunTrayProperties()
   - EnvironmentUtil.OS.RunWindowsSecurity()
   - EnvironmentUtil.OS.RunWindowSwitcher()
   - EnvironmentUtil.Shell.Desktop.CascadeWindows()
   - EnvironmentUtil.Shell.Desktop.Hide()
   - EnvironmentUtil.Shell.Desktop.Show()
   - EnvironmentUtil.Shell.Desktop.TileWindowsHorizontally()
   - EnvironmentUtil.Shell.Desktop.TileWindowsVertically()
   - EnvironmentUtil.Shell.Desktop.ToggleState()
   - EnvironmentUtil.Shell.Explorer.AddFileToRecentDocs(String)
   - EnvironmentUtil.Shell.Explorer.RefreshWindows()
   - EnvironmentUtil.Shell.StartMenu.PinItem(String)
   - EnvironmentUtil.Shell.StartMenu.UnpinItem(String)
   - EnvironmentUtil.Shell.TaskBar.Hide(Boolean)
   - EnvironmentUtil.Shell.TaskBar.PinItem(String)
   - EnvironmentUtil.Shell.TaskBar.Show(Boolean)
   - EnvironmentUtil.Shell.TaskBar.UnpinItem(String)
   - EnvironmentUtil.Theming.RemoveDesktopWallpaper()
   - EnvironmentUtil.Theming.SetDesktopWallpaper(String, EnvironmentUtil.Theming.WallpaperStyle)
   - EnvironmentUtil.Theming.SetSystemCursor(String, EnvironmentUtil.Theming.CursorType)
   - EnvironmentUtil.Theming.SetSystemVisualTheme(String, String, String)

# File-Assoc Util
Contains related Windows file association utilities. 

It can register a file extension or get system info about a registered extension.

Public Members Summary

 - Types
   - FileAssocUtil.FileExtensionInfo <Serializable>

 - Enumerations
   - FileAssocUtil.RegistryScope As Integer

 - Properties
   - FileassocUtil.FileExtensionInfo.Command As String
   - FileassocUtil.FileExtensionInfo.ContentType As String
   - FileassocUtil.FileExtensionInfo.DdeApplication As String
   - FileassocUtil.FileExtensionInfo.DdeCommand As String
   - FileassocUtil.FileExtensionInfo.DdeIfExec As String
   - FileassocUtil.FileExtensionInfo.DdeTopic As String
   - FileassocUtil.FileExtensionInfo.DefaultIcon As String
   - FileassocUtil.FileExtensionInfo.DelegateExecute As String
   - FileassocUtil.FileExtensionInfo.DropTarget As String
   - FileassocUtil.FileExtensionInfo.Executable As String
   - FileassocUtil.FileExtensionInfo.FriendlyAppName As String
   - FileassocUtil.FileExtensionInfo.FriendlyDocName As String
   - FileassocUtil.FileExtensionInfo.InfoTip As String
   - FileassocUtil.FileExtensionInfo.Name As String
   - FileassocUtil.FileExtensionInfo.NoOpen As String
   - FileassocUtil.FileExtensionInfo.QuickTip As String
   - FileassocUtil.FileExtensionInfo.ShellExtension As String
   - FileassocUtil.FileExtensionInfo.ShellNewValue As String
   - FileassocUtil.FileExtensionInfo.SupportedUriProtocols As String
   - FileassocUtil.FileExtensionInfo.TileInfo As String
   - FileassocUtil.FileExtensionInfo.Max As String <Hidden>

 - Functions
   - FileAssocUtil.GetFileExtensionInfo(String) As FileAssocUtil.FileExtensionInfo
   - FileAssocUtil.IsRegistered(String) As Boolean

 - Methods
   - FileAssocUtil.Register(FileAssocUtil.RegistryScope, String, String, Opt: String, Opt: String, Opt: Integer, Opt: String, Opt: String)

# RecycleBin Util
This class manages the system recycle bins.

It has a property to manage the master recycle bin (the one that contains the recycled files of all drives)

It can list recycled files, folders and items. Can permanently delete an specified item, or perform a cleaning operation, and more things. 

 - Child Classes
   - RecycleBinUtil.MasterBinLayout <Hidden>
   - RecycleBinUtil.Tools

 - Enumerations
   - RecycleBinUtil.CleanFlags As Integer
   - RecycleBinUtil.ItemVerbs As Integer

 - Properties
   - RecycleBinUtil.MasterBin As RecycleBinUtil.MasterBinLayout
   - RecycleBinUtil.MasterBin.KnownFolder As IKnownFolder
   - RecycleBinUtil.MasterBin.Files As IEnumerable(Of ShellFile)
   - RecycleBinUtil.MasterBin.Folders As IEnumerable(Of ShellFolder)
   - RecycleBinUtil.MasterBin.Items As IEnumerable(Of ShellObject)
   - RecycleBinUtil.MasterBin.ItemsCount As Long
   - RecycleBinUtil.MasterBin.LastRecycledFile As ShellFile
   - RecycleBinUtil.MasterBin.LastRecycledFolder As ShellFolder
   - RecycleBinUtil.MasterBin.LastRecycledItem As ShellObject
   - RecycleBinUtil.MasterBin.Size As Long

 - Functions
   - RecycleBinUtil.MasterBin.Clean(Opt: RecycleBinUtil.RecycleBinFlags) As Boolean
   - RecycleBinUtil.MasterBin.UpdateIcon() As Boolean
   - RecycleBinUtil.Tools.Clean(Char, Opt: CleanFlags) As Boolean
   - RecycleBinUtil.Tools.GetBinSize(Char) As Long
   - RecycleBinUtil.Tools.GetItemsCount(Char) As Long
   - RecycleBinUtil.Tools.GetRecycledFiles(Char) As IEnumerable(Of ShellFile)
   - RecycleBinUtil.Tools.GetRecycledFolders(Char) As IEnumerable(Of Shellfolder)
   - RecycleBinUtil.Tools.GetRecycledItems(Char) As IEnumerable(Of ShellObject)
   - RecycleBinUtil.Tools.GetLastRecycledFile(Char) As ShellFile
   - RecycleBinUtil.Tools.GetLastRecycledFolder(Char) As ShellFolder
   - RecycleBinUtil.Tools.GetLastRecycledItem(Char) As ShellObject

 - Methods
   - RecycleBinUtil.Tools.DeleteItem(ShellObject)
   - RecycleBinUtil.Tools.UndeleteItem(ShellObject)
   - RecycleBinUtil.Tools.InvokeItemVerb(ShellObject, RecycleBinUtil.ItemVerbs)
   - RecycleBinUtil.Tools.InvokeItemVerb(ShellObject, String)

# RegEdit
Contains related registry utilities.

It can do a lot of common registry operations.

Also it exposes an useful generic 'RegInfo(Of T)' class to manage registry keys or values.

Public Members Summary

 - Types
   - RegEdit.RegInfo(Of T) <Serializable>
   - RegEdit.RegInfo : Inherits RegEdit.RegInfo(Of Object) <Serializable>

 - Properties
   - RegEdit.RegInfo(Of T).RootKeyName As String
   - RegEdit.RegInfo(Of T).SubKeyPath As String
   - RegEdit.RegInfo(Of T).ValueName As String
   - RegEdit.RegInfo(Of T).ValueType As RegistryValueKind
   - RegEdit.RegInfo(Of T).ValueData As T
   - RegEdit.RegInfo(Of T).FullKeyPath As String
   - RegEdit.RegInfo(Of T).RegistryKey(Opt: RegistryKeyPermissionCheck, Opt: RegistryRights) As RegistryKey
   - RegEdit.RegInfo(Of T).RootKeyName As String
   - RegEdit.RegInfo(Of T).RootKeyName As String
   - RegEdit.RegInfo.RootKeyName As String
   - RegEdit.RegInfo.SubKeyPath As String
   - RegEdit.RegInfo.ValueName As String
   - RegEdit.RegInfo.ValueType As RegistryValueKind
   - RegEdit.RegInfo.ValueData As Object
   - RegEdit.RegInfo.FullKeyPath As String
   - RegEdit.RegInfo.RegistryKey(Opt: RegistryKeyPermissionCheck, Opt: RegistryRights) As RegistryKey
   - RegEdit.RegInfo.RootKeyName As String
   - RegEdit.RegInfo.RootKeyName As String

 - Functions
   - RegEdit.CreateSubKey(Of T)(RegInfo(Of T), RegistryKeyPermissionCheck, RegistryOptions) As RegInfo(Of T)
   - RegEdit.CreateSubKey(Of T)(String, RegistryKeyPermissionCheck, RegistryOptions) As RegInfo(Of T)
   - RegEdit.CreateSubKey(Of T)(String, String, RegistryKeyPermissionCheck, RegistryOptions) As RegInfo(Of T)
   - RegEdit.CreateSubKey(RegInfo, RegistryKeyPermissionCheck, RegistryOptions) As RegistryKey
   - RegEdit.CreateSubKey(String, RegistryKeyPermissionCheck, RegistryOptions) As RegistryKey
   - RegEdit.CreateSubKey(String, String, RegistryKeyPermissionCheck, RegistryOptions) As RegistryKey
   - RegEdit.ExistSubKey(String) As Boolean
   - RegEdit.ExistSubKey(String, String) As Boolean
   - RegEdit.ExistValue(String, String) As Boolean
   - RegEdit.ExistValue(String, String, String) As Boolean
   - RegEdit.ExportKey(String, String) As Boolean
   - RegEdit.ExportKey(String, String, String) As Boolean
   - RegEdit.FindSubKey(String, String, Boolean, Boolean, SearchOption) As IEnumerable(Of RegInfo)
   - RegEdit.FindSubKey(String, String, String, Boolean, Boolean, SearchOption) As IEnumerable(Of RegInfo)
   - RegEdit.FindValue(String, String, Boolean, Boolean, SearchOption) As IEnumerable(Of RegInfo)
   - RegEdit.FindValue(String, String, String, Boolean, Boolean, SearchOption) As IEnumerable(Of RegInfo)
   - RegEdit.FindValueData(String, String, String, Boolean, Boolean, SearchOption) As IEnumerable(Of RegInfo)
   - RegEdit.GetRootKey(String) As RegistryKey
   - RegEdit.GetRootKeyName(String) As String
   - RegEdit.GetSubKeyPath(String) As String
   - RegEdit.GetValueData(Of T)(RegInfo(Of T), RegistryValueOptions) As T
   - RegEdit.GetValueData(Of T)(String, String, RegistryValueOptions) As T
   - RegEdit.GetValueData(Of T)(String, String, String, RegistryValueOptions) As T
   - RegEdit.GetValueData(RegInfo, RegistryValueOptions) As Object
   - RegEdit.GetValueData(String, String, RegistryValueOptions) As Object
   - RegEdit.GetValueData(String, String, String, RegistryValueOptions) As Object
   - RegEdit.ImportRegFile(String) As Boolean
   - RegEdit.ValueIsEmpty(String, String) As Boolean
   - RegEdit.ValueIsEmpty(String, String, String) As Boolean

 - Methods
   - RegEdit.CopyKeyTree(String, String)
   - RegEdit.CopyKeyTree(String, String, String, String)
   - RegEdit.CopySubKeys(RegistryKey, RegistryKey)
   - RegEdit.CopySubKeys(String, String)
   - RegEdit.CopySubKeys(String, String, String, String)
   - RegEdit.CopyValue(String, String, String, String)
   - RegEdit.CopyValue(String, String, String, String, String, String)
   - RegEdit.CreateValue(Of T)(RegInfo(Of T))
   - RegEdit.CreateValue(Of T)(String, String, String, T, RegistryValueKind)
   - RegEdit.CreateValue(Of T)(String, String, T, RegistryValueKind)
   - RegEdit.DeleteSubKey(Of T)(RegInfo(Of T), Boolean)
   - RegEdit.DeleteSubKey(String, Boolean)
   - RegEdit.DeleteSubKey(String, String, Boolean)
   - RegEdit.DeleteValue(Of T)(RegInfo(Of T), Boolean)
   - RegEdit.DeleteValue(String, String, Boolean)
   - RegEdit.DeleteValue(String, String, String, Boolean)
   - RegEdit.JumpToKey(String)
   - RegEdit.JumpToKey(String, String)
   - RegEdit.MoveKeyTree(String, String)
   - RegEdit.MoveKeyTree(String, String, String, String)
   - RegEdit.MoveSubKeys(String, String)
   - RegEdit.MoveSubKeys(String, String, String, String)
   - RegEdit.MoveValue(String, String, String, String)
   - RegEdit.MoveValue(String, String, String, String, String, String)

# System Restarter
Safely shutdowns, restarts or logoffs the local or remote computer.

A shutdown operation can be programmed and/or aborted at any time.

 - Constants
   - SystemRestarter.MaxShutdownTimeout As Integer

 - Enumerations
   - SystemRestarter.LogOffMode As UInteger
   - SystemRestarter.ShutdownMode As UInteger
   - SystemRestarter.ShutdownReason As UInteger <Flags>
   - SystemRestarter.ShutdownPlanning As UInteger

 - Functions
   - SystemRestarter.Abort(Opt: String, Opt: Boolean) As Boolean
   - SystemRestarter.LogOff(Opt: LogOffMode, Opt: ShutdownReason, Opt: Boolean) As Boolean
   - SystemRestarter.PowerOff(Opt: String, Opt: Integer, Opt: String, Opt: ShutdownMode, Opt: ShutdownReason, Opt: ShutdownPlanning, Opt: Boolean) As Boolean
   - SystemRestarter.Restart(Opt: String, Opt: Integer, Opt: String, Opt: ShutdownMode, Opt: ShutdownReason, Opt: ShutdownPlanning, Opt: Boolean) As Boolean
   - SystemRestarter.RestartApps(Opt: String, Opt: Integer, Opt: String, Opt: ShutdownMode, Opt: ShutdownReason, Opt: ShutdownPlanning, Opt: Boolean) As Boolean
   - SystemRestarter.Shutdown(Opt: String, Opt: Integer, Opt: String, Opt: ShutdownMode, Opt: ShutdownReason, Opt: ShutdownPlanning, Opt: Boolean) As Boolean
   - SystemRestarter.HybridShutdown(Opt: String, Opt: Integer, Opt: String, Opt: ShutdownMode, Opt: ShutdownReason, Opt: ShutdownPlanning, Opt: Boolean) As Boolean

# User-Account Util
Contains related Windows user-account utilities.

It can add, delete or find an user, and much more.

Public Members Summary

 - Properties
   - UserAccountUtil.CurrentUser As UserPrincipal
   - UserAccountUtil.CurrentUserIsAdmin As Boolean

 - Functions
   - UserAccountUtil.Create(String, String, String, String, Boolean, Boolean) As UserPrincipal
   - UserAccountUtil.FindProfilePath(SecurityIdentifier) As String
   - UserAccountUtil.FindProfilePath(String) As String
   - UserAccountUtil.FindSid(String) As SecurityIdentifier
   - UserAccountUtil.FindUser(SecurityIdentifier) As UserPrincipal
   - UserAccountUtil.FindUser(String) As UserPrincipal
   - UserAccountUtil.FindUsername(SecurityIdentifier) As String
   - UserAccountUtil.GetAllUsers() As List(Of UserPrincipal)
   - UserAccountUtil.IsAdmin(String) As Boolean
   - UserAccountUtil.IsMemberOfGroup(String, String) As Boolean
   - UserAccountUtil.IsMemberOfGroup(String, WellKnownSidType) As Boolean

 - Methods
   - UserAccountUtil.Add(String, String, String, String, Boolean, Boolean, WellKnownSidType)
   - UserAccountUtil.Add(UserPrincipal, WellKnownSidType)
   - UserAccountUtil.Delete(String)
