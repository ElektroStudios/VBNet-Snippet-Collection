# System snippets category
These snippets are intended to help managing the (operating)system in some aspects.

# General Information about this repository
 - .snippet files contains the formatted code that can be loaded through Visual Studio's code editor context menu.
 - .vb files contains the raw code that can be copied then pasted in any project.
 - Each .snippet and .vb file contains a #Region section and/or Xml documentation with code examples.
 
Feel free to use and/or modify any file of this repository.

If you like the job I've done, then please contribute with improvements to these snippets or by adding new ones.

#Environment Variable Util
Contains related Environment Variable utilities, such as register or unregister a variable.

Public Members Summary

 - Types
   - EnvironmentVariableUtil.VariableInfo <Serializable>

 - Constructors
   - EnvironmentVariableUtil.VariableInfo.New(EnvironmentVariableUtil.VariableScope, String, String)

 - Enumerations
   - EnvironmentVariableUtil.VariableScope As Integer

 - Properties
   - EnvironmentVariableUtil.CurrentVariables(EnvironmentVariableUtil.VariableScope) As ReadOnlyCollection(Of EnvironmentVariableUtil.VariableInfo)
   - EnvironmentVariableUtil.VariableInfo.Name As String
   - EnvironmentVariableUtil.VariableInfo.Value As String
   - EnvironmentVariableUtil.VariableInfo.Scope As EnvironmentVariableUtil.VariableScope

 - Functions
   - EnvironmentVariableUtil.GetValue(EnvironmentVariableUtil.VariableScope, String, Boolean) As Opt: Boolean
   - EnvironmentVariableUtil.GetVariableInfo(EnvironmentVariableUtil.VariableScope, String, Opt: Boolean) As EnvironmentVariableUtil.VariableInfo

 - Methods
   - EnvironmentVariableUtil.RegisterVariable(EnvironmentVariableUtil.VariableInfo, Opt: Boolean)
   - EnvironmentVariableUtil.RegisterVariable(EnvironmentVariableUtil.VariableScope, String, String, Opt: Boolean)
   - EnvironmentVariableUtil.UnregisterVariable(EnvironmentVariableUtil.VariableScope, String, Opt: Boolean)
   - EnvironmentVariableUtil.UnregisterVariable(EnvironmentVariableUtil.VariableInfo, Opt: Boolean)

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

#FileSystem Util
Contains related FileSystem utilities, such as filepath validation or invoke an item verb.

Public Members Summary

 - Functions
   - FileSystemUtil.GetItemVerbs(String) As IEnumerable(Of FolderItemVerb)
   - FileSystemUtil.ItemNameIsInvalid(String) As Boolean
   - FileSystemUtil.ItemNameOrPathIsInvalid(String) As Boolean
   - FileSystemUtil.ItemPathIsInvalid(String) As Boolean

 - Methods
   - FileSystemUtil.InvokeItemVerb(String, String)

#Power Util
Contains related system powering utilities.

Public Members Summary

 - Functions
   - PowerUtil.SendAwakeSignal As Boolean
   - PowerUtil.RemoveAwakeSignal As Boolean

# RecycleBin Util
This class manages the system recycle bins.

It has a property to manage the master recycle bin (the one that contains the recycled files of all drives)

It can list recycled files, folders and items. Can permanently delete an specified item, or perform a cleaning operation, and more things. 

Public Members Summary

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

# RegEdit (Registry Editor)
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

#Shell Util
Contains related Windows shell utilities, such as pin or unpin items in taskbar and startmenu, show/hide taskbar and startmenu, get all the Explorer windows instances, refresh Explorer windows, or add a file to recent docs.

Public Members Summary

 - Child Classes
   - ShellUtil.Applets
   - ShellUtil.Desktop
   - ShellUtil.Explorer
   - ShellUtil.StartMenu
   - ShellUtil.TaskBar

 - Enumerations
   - ShellUtil.TaskBar.TaskbarVisibility As Integer

 - Properties
   - ShellUtil.Explorer.ExplorerWindows As ReadOnlyCollection(Of ShellBrowserWindow)
   - ShellUtil.Explorer.ExplorerWindowsFolders As ReadOnlyCollection(Of Shell32.Folder2)
   - ShellUtil.TaskBar.ClassName() As String
   - ShellUtil.TaskBar.Hwnd() As Intptr

 - Methods
   - ShellUtil.Applets.RunDateTime()
   - ShellUtil.Applets.RunExecuteDialog()
   - ShellUtil.Applets.RunFindComputer()
   - ShellUtil.Applets.RunFindFiles()
   - ShellUtil.Applets.RunFindPrinter()
   - ShellUtil.Applets.RunHelpCenter()
   - ShellUtil.Applets.RunSearchCommand()
   - ShellUtil.Applets.RunTrayProperties()
   - ShellUtil.Applets.RunWindowsSecurity()
   - ShellUtil.Applets.RunWindowSwitcher()
   - ShellUtil.Desktop.CascadeWindows()
   - ShellUtil.Desktop.Hide()
   - ShellUtil.Desktop.Show()
   - ShellUtil.Desktop.TileWindowsHorizontally()
   - ShellUtil.Desktop.TileWindowsVertically()
   - ShellUtil.Desktop.ToggleState()
   - ShellUtil.Explorer.AddFileToRecentDocs(String)
   - ShellUtil.Explorer.RefreshWindows()
   - ShellUtil.StartMenu.PinItem(String)
   - ShellUtil.StartMenu.UnpinItem(String)
   - ShellUtil.TaskBar.Hide(Boolean)
   - ShellUtil.TaskBar.PinItem(String)
   - ShellUtil.TaskBar.Show(Boolean)
   - ShellUtil.TaskBar.UnpinItem(String)
   
#System Info
Contains related Windows operating system's information utilities.

Public Members Summary

 - Child Classes

   - SystemInfo.OS
   - SystemInfo.Programs

 - Enumerations
   - SystemInfo.Architecture As Integer

 - Properties
   - SystemInfo.OS.CurrentArchitecture As SystemInfo.Architecture
   - SystemInfo.Programs.DefaultWebBrowser() As String
   - SystemInfo.Programs.IExplorerVersion() As Version

#System Notifier
Notifies the system about changes on the environment to update quickly

Public Members Summary

 - Methods
   - SystemNotifier.NotifyDirectoryAttributesChanged(String)
   - SystemNotifier.NotifyDirectoryCreated(String)
   - SystemNotifier.NotifyDirectoryDeleted(String)
   - SystemNotifier.NotifyDirectoryRenamed(String, String)
   - SystemNotifier.NotifyDriveAdded(String, Boolean)
   - SystemNotifier.NotifyDriveRemoved(String)
   - SystemNotifier.NotifyFileAssociationChanged()
   - SystemNotifier.NotifyFileAttributesChanged(String)
   - SystemNotifier.NotifyFileCreated(String)
   - SystemNotifier.NotifyFileDeleted(String)
   - SystemNotifier.NotifyFileRenamed(String, String)
   - SystemNotifier.NotifyFreespaceChanged(String)
   - SystemNotifier.NotifyMediaInserted(String)
   - SystemNotifier.NotifyMediaRemoved(String)
   - SystemNotifier.NotifyNetworkFolderShared(String)
   - SystemNotifier.NotifyNetworkFolderUnshared(String)
   - SystemNotifier.NotifyUpdateDirectory(String)
   - SystemNotifier.NotifyUpdateImage()
   - SystemNotifier.ReloadSystemCursors()
   - SystemNotifier.ReloadSystemIcons()

# System Restarter
Safely shutdowns, restarts or logoffs the local or remote computer.

A shutdown operation can be programmed and/or aborted at any time.

Public Members Summary

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

#Theming Util
Contains related theming/personalization utilities, such as set the current system cursors, theme or wallpaper.

Public Members Summary

 - Child Classes
   - ThemingUtil.Cursors
   - ThemingUtil.Fonts
   - ThemingUtil.ScreenSavers
   - ThemingUtil.VisualThemes
   - ThemingUtil.Wallpapers

 - Types
   - ThemingUtil.VisualThemes.ThemeInfo

 - Constructors
   - ThemingUtil.VisualThemes.ThemeInfo.New(String, String, String)

 - Enumerations
   - ThemingUtil.Cursors.CursorType
   - ThemingUtil.Wallpapers.WallpaperStyle

 - Properties
   - ThemingUtil.Fonts.InstalledFonts As IEnumerable(Of FontFamily)
   - ThemingUtil.Fonts.InstalledFontNames As IEnumerable(Of String)
   - ThemingUtil.ScreenSavers.CurrentScreensaver As String
   - ThemingUtil.VisualThemes.AeroEnabled() As Boolean
   - ThemingUtil.VisualThemes.AeroSupported() As Boolean
   - ThemingUtil.VisualThemes.CurrentTheme() As ThemingUtil.VisualThemes.ThemeInfo
   - ThemingUtil.Wallpapers.CurrentWallpaper() As String
   - ThemingUtil.Wallpapers.WallpaperAsJpegIsSupported() As Boolean
   - ThemingUtil.Wallpapers.WallpaperStylesFitFillAreSupported() As Boolean

 - Functions
   - ThemingUtil.Fonts.IsFontInstalled(String) As Boolean
   - ThemingUtil.Fonts.IsFontInstalled(FontFamily) As Boolean

 - Methods
   - ThemingUtil.Cursors.SetSystemCursor(String, ThemingUtil.Cursors.CursorType)
   - ThemingUtil.ScreenSavers.SetScreensaver(String)
   - ThemingUtil.VisualThemes.SetSystemVisualTheme(String, String, String)
   - ThemingUtil.Wallpapers.RemoveDesktopWallpaper()
   - ThemingUtil.Wallpapers.SetDesktopWallpaper(String, ThemingUtil.Wallpapers.WallpaperStyle)

#Tweaking Util
Contains related Windows tweaking utilities to change system parameters.

Public Members Summary

 - Child Classes
   - TweakingUtil.SystemParameters

 - Properties
   - TweakingUtil.SystemParameters.ActiveWindowTrackingEnabled As Boolean
   - TweakingUtil.SystemParameters.ActiveWindowTrackingTimeout As UShort
   - TweakingUtil.SystemParameters.BeepEnabled As Boolean
   - TweakingUtil.SystemParameters.BlockSendInputResetsEnabled As Boolean
   - TweakingUtil.SystemParameters.BorderMultiplierFactor As Integer
   - TweakingUtil.SystemParameters.CaretWidth As Integer
   - TweakingUtil.SystemParameters.CleartypeEnabled As Boolean
   - TweakingUtil.SystemParameters.ClientAreaAnimationEnabled As Boolean
   - TweakingUtil.SystemParameters.ComboBoxAnimationEnabled As Boolean
   - TweakingUtil.SystemParameters.CursorShadowEnabled As Boolean
   - TweakingUtil.SystemParameters.DoubleClickSize As Size
   - TweakingUtil.SystemParameters.DoubleClickTime As Integer
   - TweakingUtil.SystemParameters.DragFullWindowsEnabled As Boolean
   - TweakingUtil.SystemParameters.DragSize As Size
   - TweakingUtil.SystemParameters.DropShadowEnabled As Boolean
   - TweakingUtil.SystemParameters.FlatMenuEnabled As Boolean
   - TweakingUtil.SystemParameters.FocusBorderSize As Size
   - TweakingUtil.SystemParameters.FontSmoothingContrast As Integer
   - TweakingUtil.SystemParameters.FontSmoothingEnabled As Boolean
   - TweakingUtil.SystemParameters.ForegroundFlashCount As UShort
   - TweakingUtil.SystemParameters.ForegroundLockTimeout As UShort
   - TweakingUtil.SystemParameters.HotTrackingEnabled As Boolean
   - TweakingUtil.SystemParameters.HungAppTimeout As Integer
   - TweakingUtil.SystemParameters.IconSpacing As Size
   - TweakingUtil.SystemParameters.IconTitleWrappingEnabled As Boolean
   - TweakingUtil.SystemParameters.KeyboardDelay As Integer
   - TweakingUtil.SystemParameters.KeyboardSpeed As Integer
   - TweakingUtil.SystemParameters.ListBoxSmoothScrollingEnabled As Boolean
   - TweakingUtil.SystemParameters.MenuAccessKeysUnderlined As Boolean
   - TweakingUtil.SystemParameters.MenuAnimationEnabled As Boolean
   - TweakingUtil.SystemParameters.MenuFadeEnabled As Boolean
   - TweakingUtil.SystemParameters.MenuShowDelay As Integer
   - TweakingUtil.SystemParameters.MessageDuration As Long
   - TweakingUtil.SystemParameters.MouseButtonsSwapEnabled As Boolean
   - TweakingUtil.SystemParameters.MouseClickLockEnabled As Boolean
   - TweakingUtil.SystemParameters.MouseClickLockTime As Integer
   - TweakingUtil.SystemParameters.MouseHoverSize As Size
   - TweakingUtil.SystemParameters.MouseHoverTime As Integer
   - TweakingUtil.SystemParameters.MouseSonarEnabled As Boolean
   - TweakingUtil.SystemParameters.MouseSpeed As Integer
   - TweakingUtil.SystemParameters.MouseTrailAmount As Integer
   - TweakingUtil.SystemParameters.MouseVanishEnabled As Boolean
   - TweakingUtil.SystemParameters.MouseWheelScrollLines As Integer
   - TweakingUtil.SystemParameters.OverlappedContentEnabled As Boolean
   - TweakingUtil.SystemParameters.PopupMenuAlignment As LeftRightAlignment
   - TweakingUtil.SystemParameters.ScreensaverEnabled As Boolean
   - TweakingUtil.SystemParameters.ScreensaverPath As String
   - TweakingUtil.SystemParameters.ScreensaverTimeout As Integer
   - TweakingUtil.SystemParameters.ScreensaveSecureEnabled As Boolean
   - TweakingUtil.SystemParameters.SelectionFadeEnabled As Boolean
   - TweakingUtil.SystemParameters.SnapToDefaultEnabled As Boolean
   - TweakingUtil.SystemParameters.SystemDateTime As Date
   - TweakingUtil.SystemParameters.SystemLanguageBarEnabled As Boolean
   - TweakingUtil.SystemParameters.TitleBarGradientEnabled As Boolean
   - TweakingUtil.SystemParameters.ToolTipAnimationEnabled As Boolean
   - TweakingUtil.SystemParameters.UIEffectsEnabled As Boolean
   - TweakingUtil.SystemParameters.WaitToKillAppTimeout As Integer
   - TweakingUtil.SystemParameters.WaitToKillServiceTimeout As Integer
   - TweakingUtil.SystemParameters.WheelscrollChars As Integer

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
