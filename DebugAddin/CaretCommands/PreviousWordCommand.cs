using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using EnvDTE80;
using Task = System.Threading.Tasks.Task;
using Microsoft.VisualStudio.TextManager.Interop;

namespace DebugAddin.CaretCommands
  {
  internal sealed class PreviousWordCommand
    {
    public const int CommandId = 0x0200;
    public static readonly Guid CommandSet = new Guid("925c9a20-7c34-4406-b9a8-17a1b657d1d1");
    private readonly AsyncPackage package;

    private PreviousWordCommand(AsyncPackage package, OleMenuCommandService commandService)
      {
      this.package = package ?? throw new ArgumentNullException(nameof(package));
      commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

      var menuCommandID = new CommandID(CommandSet, CommandId);
      var menuItem = new MenuCommand(this.Execute, menuCommandID);
      commandService.AddCommand(menuItem);
      }

    public static PreviousWordCommand Instance
      {
      get;
      private set;
      }

    private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
      {
      get
        {
        return this.package;
        }
      }

    public static async Task InitializeAsync(AsyncPackage package)
      {
      // Switch to the main thread - the call to AddCommand in CaretCommands's constructor requires
      // the UI thread.
      await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

      OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
      Instance = new PreviousWordCommand(package, commandService);
      }

    DTE2 dte = (DTE2)Package.GetGlobalService(typeof(DTE));
    IVsTextManager textManager = (IVsTextManager)Package.GetGlobalService(typeof(SVsTextManager));

    private void Execute(object sender, EventArgs e)
      {
      IVsTextView textView;
      textManager.GetActiveView(1, null, out textView);
      
      int lineOld, columnOld;
      textView.GetSelection(out _, out _, out lineOld, out columnOld);

      if (columnOld == 0 && lineOld > 0)
        {
        string line;
        textView.GetTextStream(lineOld - 1, 0, lineOld, 0, out line);

        // move cursor to the last non-whitespace char in the previous line
        for (int i = line.Length - 2; i >= columnOld; --i)
          if (line[i] != ' ')
            {
            textView.SetSelection(lineOld - 1, i + 1, lineOld - 1, i + 1);
            return;
            }
        }
      dte.ExecuteCommand("Edit.WordPrevious");
      }
    }
  }
