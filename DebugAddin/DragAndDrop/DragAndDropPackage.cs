//------------------------------------------------------------------------------
// <copyright file="DragAndDropPackage.cs" company="Microsoft">
//     Copyright (c) Microsoft.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;

namespace DebugAddin
  {
  [PackageRegistration(UseManagedResourcesOnly = true)]
  [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
  [Guid(DragAndDropPackage.PackageGuidString)]
  [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]
  [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
  [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
  public sealed class DragAndDropPackage : Package
    {
    public const string PackageGuidString = "7f0d05e6-785c-4b2e-ae98-69df2a911976";
    int hHook = 0;

#pragma warning disable 618
    public DragAndDropPackage()
    {
      messageId = RegisterWindowMessage("MyDragDropMessage32");
      // Create an instance of HookProc.
      hook = new HookProc(MouseHookProc);
      hHook = SetWindowsHookEx(4, hook, (IntPtr)0, AppDomain.GetCurrentThreadId());
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern int SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hInstance, int threadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern bool UnhookWindowsHookEx(int idHook);

    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern int CallNextHookEx(int idHook, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern uint RegisterWindowMessage(string lpString);

    public delegate int HookProc(int nCode, IntPtr wParam, IntPtr lParam);
    static HookProc hook;
    static uint messageId = 0;

    //Declare the wrapper managed MouseHookStruct class.
    [StructLayout(LayoutKind.Sequential)]
    public class tagCWPSTRUCT
    {
      public IntPtr lParam;
      public IntPtr wParam;
      public uint message;
      // HWND hwnd;
    }

    public static int MouseHookProc(int nCode, IntPtr wParam, IntPtr lParam)
    {
      ThreadHelper.ThrowIfNotOnUIThread();
      if (nCode < 0)
        return CallNextHookEx(0, nCode, wParam, lParam);

      tagCWPSTRUCT inputParams = (tagCWPSTRUCT)Marshal.PtrToStructure(lParam, typeof(tagCWPSTRUCT));

      if (inputParams.message == messageId)
      {
        DTE2 dte = (DTE2)ServiceProvider.GlobalProvider.GetService(typeof(DTE));
        Assumes.Present(dte);
        dte.ItemOperations.OpenFile(File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + "devenv.exe.buffer"));
      }

      return CallNextHookEx(0, nCode, wParam, lParam);
    }

    #region Package Members

    protected override void Initialize()
      {
        base.Initialize();
      }

    protected override void Dispose(bool disposing)
      {
        UnhookWindowsHookEx(hHook);
      }

    #endregion
    }
  }
