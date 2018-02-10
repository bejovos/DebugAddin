using System;
using EnvDTE;
using EnvDTE80;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TeamFoundation.VersionControl;
using Microsoft.TeamFoundation.VersionControl.Client;

namespace DebugAddin.ViewHistoryCommand
  {
  internal sealed class ViewHistoryCommand
    {
    public const int CommandId = 0x0100;

    public static readonly Guid CommandSet = new Guid("c2280335-99cc-4f01-8b7e-ac375864d19a");

    private readonly Package package;

    private ViewHistoryCommand(Package package)
      {
      if (package == null)
        {
        throw new ArgumentNullException("package");
        }

      this.package = package;

      OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
      if (commandService != null)
        {
        var menuCommandID = new CommandID(CommandSet, CommandId);
        var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
        commandService.AddCommand(menuItem);
        }
      }

    public static ViewHistoryCommand Instance
      {
      get;
      private set;
      }

    private IServiceProvider ServiceProvider
      {
      get
        {
        return this.package;
        }
      }

    public static void Initialize(Package package)
      {
      Instance = new ViewHistoryCommand(package);
      }

    private void MenuItemCallback(object sender, EventArgs e)
      {
      DTE2 dte = (DTE2)Package.GetGlobalService(typeof(DTE));
      VersionControlExt versionControlExt = dte.GetObject("Microsoft.VisualStudio.TeamFoundation.VersionControl.VersionControlExt") as VersionControlExt;
      versionControlExt.History.Show(dte.ActiveDocument.FullName, 
        VersionSpec.Latest, 0, RecursionType.None);
      }
    }
  }
