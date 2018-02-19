using System;
using System.ComponentModel.Design;
using System.Globalization;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace DebugAddin.BuildStartupProject
  {
  internal sealed class BuildStartupProjectCommand
    {
    public const int CommandId = 0x0100;

    public static readonly Guid CommandSet = new Guid("8d15eab3-f107-4d6a-b4c3-59ba5b0aadde");

    private readonly Package package;

    private BuildStartupProjectCommand(Package package)
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

    public static BuildStartupProjectCommand Instance
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
      Instance = new BuildStartupProjectCommand(package);
      }

    private void MenuItemCallback(object sender, EventArgs e)
      {
      try
        {
        DTE2 dte = (DTE2)Package.GetGlobalService(typeof(DTE));        
        string project_name = "";
        foreach (String s in (Array)dte.Solution.SolutionBuild.StartupProjects)
          {
          project_name += s;
          }

        Project startupProject = null;
        foreach (Project p in CmdArgs.Utils.GetAllProjectsInSolution())
          {
          if (p.UniqueName == project_name)
            {
            startupProject= p;
            break;
            }
          }
        UIHierarchyItem item = CmdArgs.Utils.FindUIHierarchyItem(startupProject);
        item.Select(vsUISelectionType.vsUISelectionTypeSelect);
        dte.ToolWindows.SolutionExplorer.Parent.Activate();
        dte.ExecuteCommand("Build.BuildSelection");

        foreach (EnvDTE.Window window in dte.Windows)
          {
          if (window.Caption.StartsWith("Pending Changes"))
            {
            window.Activate();
            break;
            }
          }
        }
      catch (Exception ex)
        {
        CmdArgs.Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace);
        }
      }
    }
  }
