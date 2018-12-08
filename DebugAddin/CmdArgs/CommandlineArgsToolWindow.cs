//------------------------------------------------------------------------------
// <copyright file="CommandlineArgsToolWindow.cs" company="Microsoft">
//     Copyright (c) Microsoft.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace DebugAddin.CmdArgs
  {
  using System;
  using System.Runtime.InteropServices;
  using Microsoft.VisualStudio.Shell;

  /// <summary>
  /// This class implements the tool window exposed by this package and hosts a user control.
  /// </summary>
  /// <remarks>
  /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane,
  /// usually implemented by the package implementer.
  /// <para>
  /// This class derives from the ToolWindowPane class provided from the MPF in order to use its
  /// implementation of the IVsUIElementPane interface.
  /// </para>
  /// </remarks>
  [Guid("767ce98b-8f7b-41ee-b33f-750b2a6e2fcb")]
  public class CommandlineArgsToolWindow : ToolWindowPane
    {
    CommandlineArgsToolWindowControl myControl;
    /// <summary>
    /// Initializes a new instance of the <see cref="CommandlineArgsToolWindow"/> class.
    /// </summary>
    public CommandlineArgsToolWindow() : base(null)
      {
      this.Caption = "Commandline Arguments";

      // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
      // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
      // the object returned by the Content property.
      myControl = new CommandlineArgsToolWindowControl();
      this.Content = myControl;
      }

    public override void OnToolWindowCreated()
      {
      myControl.toolwindow = this;
      }

    protected override void OnClose()
      {
      myControl.toolwindow = null;
      }
    }
  }
