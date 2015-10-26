' ***********************************************************************
' Author   : Elektro
' Modified : 26-October-2015
' ***********************************************************************
' <copyright file="IPC Util.vb" company="Elektro Studios">
'     Copyright (c) Elektro Studios. All rights reserved.
' </copyright>
' ***********************************************************************

#Region " Required assembly references "

' Microsoft Windows UI Automation: 
' UIAutomationClient (UIAutomationClient.dll)
' UIAutomationTypes  (UIAutomationTypes.dll)

#End Region

#Region " Public Members Summary "

#Region " Functions "

' IpcUtil.GetTitlebarText(IntPtr) As String

#End Region

#End Region

#Region " Option Statements "

Option Strict On
Option Explicit On
Option Infer Off

#End Region

#Region " Imports "

Imports System
'Imports System.Runtime.InteropServices
'Imports System.Text
'Imports System.ComponentModel
'Imports System.Linq.Expressions
'Imports System.Reflection
Imports System.Windows.Automation

#End Region

#Region " IpcUtil "

''' ----------------------------------------------------------------------------------------------------
''' <summary>
''' Contains related Inter-process communication (IPC) utilities.
''' </summary>
''' ----------------------------------------------------------------------------------------------------
Public Module IpcUtil

#Region " P/Invoking "

    '    ''' ----------------------------------------------------------------------------------------------------
    '    ''' <summary>
    '    ''' Platform Invocation methods (P/Invoke), access unmanaged code.
    '    ''' This class does not suppress stack walks for unmanaged code permission.
    '    ''' <see cref="System.Security.SuppressUnmanagedCodeSecurityAttribute"/>  must not be applied to this class.
    '    ''' This class is for methods that can be used anywhere because a stack walk will be performed.
    '    ''' </summary>
    '    ''' ----------------------------------------------------------------------------------------------------
    '    ''' <remarks>
    '    ''' <see href="http://msdn.microsoft.com/en-us/library/ms182161.aspx"/>
    '    ''' </remarks>
    '    ''' ----------------------------------------------------------------------------------------------------
    '    Private NotInheritable Class NativeMethods
    '
    '#Region " Functions "
    '
    '#End Region
    '
    '    End Class

#End Region

#Region " Public Methods "

    ''' ----------------------------------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the titlebar's text of the specified window.
    ''' </summary>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <example> This is a code example.
    ''' <code>
    ''' MessageBox.Show(IpcUtil.GetTitlebarText(Process.GetCurrentProcess.MainWindowHandle))
    ''' MessageBox.Show(IpcUtil.GetTitlebarText(Process.GetProcessesByName("notepad").First.MainWindowHandle))
    ''' </code>
    ''' </example>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <param name="hWnd">
    ''' An <see cref="IntPtr"/> to the window.
    ''' </param>
    ''' ----------------------------------------------------------------------------------------------------
    ''' <returns>
    ''' The titlebar's text of the target window.
    ''' </returns>
    ''' ----------------------------------------------------------------------------------------------------
    <DebuggerStepThrough>
    Public Function GetTitlebarText(ByVal hWnd As IntPtr) As String

        Dim window As AutomationElement = AutomationElement.FromHandle(hWnd)
        Dim condition As New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TitleBar)
        Dim titleBar As AutomationElement = window.FindFirst(TreeScope.Children, condition)

        Return titleBar.Current.Name

    End Function

#End Region

End Module

#End Region
