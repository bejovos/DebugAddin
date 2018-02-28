using System;
using System.ComponentModel.Design;
using System.Globalization;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace DebugAddin.DiffOpenDocument
  {
  /// <summary>
  /// Command handler
  /// </summary>
  internal sealed class DiffOpenDocumentCommand
    {
    /// <summary>
    /// Command ID.
    /// </summary>
    public const int CommandId = 0x0100;

    /// <summary>
    /// Command menu group (command set GUID).
    /// </summary>
    public static readonly Guid CommandSet = new Guid("c9b8c574-1f8b-40c7-b1f2-fc524f5497cf");

    /// <summary>
    /// VS Package that provides this command, not null.
    /// </summary>
    private readonly Package package;

    /// <summary>
    /// Initializes a new instance of the <see cref="DiffOpenDocumentCommand"/> class.
    /// Adds our command handlers for menu (commands must exist in the command table file)
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    private DiffOpenDocumentCommand(Package package)
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

    /// <summary>
    /// Gets the instance of the command.
    /// </summary>
    public static DiffOpenDocumentCommand Instance
      {
      get;
      private set;
      }

    /// <summary>
    /// Gets the service provider from the owner package.
    /// </summary>
    private IServiceProvider ServiceProvider
      {
      get
        {
        return this.package;
        }
      }

    /// <summary>
    /// Initializes the singleton instance of the command.
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    public static void Initialize(Package package)
      {
      Instance = new DiffOpenDocumentCommand(package);
      }

    /// <summary>
    /// This function is the callback used to execute the command when the menu item is clicked.
    /// See the constructor to see how the menu item is associated with this function using
    /// OleMenuCommandService service and MenuCommand class.
    /// </summary>
    /// <param name="sender">Event sender.</param>
    /// <param name="e">Event args.</param>
    private void MenuItemCallback(object sender, EventArgs e)
      {
      DTE2 dte = (DTE2)Package.GetGlobalService(typeof(DTE));
      
      string fullname = dte.ActiveDocument.FullName;
      int lineNumber = (dte.ActiveDocument.Selection as TextSelection).ActivePoint.Line;
      dte.ActiveDocument.Close();
      dte.ItemOperations.OpenFile(fullname);
      dte.ExecuteCommand("Edit.GoTo " + lineNumber);
      }
    }
  }
