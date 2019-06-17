namespace DebugAddin.CmdArgsToolWindow
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
  [Guid("8144752c-d149-45d9-b919-0fdd3d88de1e")]
  public class CmdArgsToolWindow : ToolWindowPane
    {
    /// <summary>
    /// Initializes a new instance of the <see cref="CmdArgsToolWindow"/> class.
    /// </summary>
    public CmdArgsToolWindow() : base(null)
      {
      this.Caption = "Commandline Arguments";

      // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
      // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
      // the object returned by the Content property.
      this.Content = new CmdArgsToolWindowControl();
      }

    public override void OnToolWindowCreated()
      {
      (this.Content as CmdArgsToolWindowControl).toolwindow = this;
      }

    protected override void OnClose()
      {
      (this.Content as CmdArgsToolWindowControl).toolwindow = null;
      }
    }
  }
