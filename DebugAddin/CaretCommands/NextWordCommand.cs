using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using EnvDTE80;
using Task = System.Threading.Tasks.Task;
using Microsoft.VisualStudio.TextManager.Interop;

namespace DebugAddin.CaretCommands
  {
  internal sealed class NextWordCommand
    {
    public const int CommandId = 0x0100;
    public static readonly Guid CommandSet = new Guid("925c9a20-7c34-4406-b9a8-17a1b657d1d1");
    private readonly AsyncPackage package;

    private NextWordCommand(AsyncPackage package, OleMenuCommandService commandService)
      {
      this.package = package ?? throw new ArgumentNullException(nameof(package));
      commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

      var menuCommandID = new CommandID(CommandSet, CommandId);
      var menuItem = new MenuCommand(this.Execute, menuCommandID);
      commandService.AddCommand(menuItem);
      }

    public static NextWordCommand Instance
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
      Instance = new NextWordCommand(package, commandService);
      }

    DTE2 dte = (DTE2)Package.GetGlobalService(typeof(DTE));
    IVsTextManager textManager = (IVsTextManager)Package.GetGlobalService(typeof(SVsTextManager));

    private int GetLastPosition(string line)
      {
      if (line != null)
        for (int i = line.Length - 1; i >= 0; --i)
          if (line[i] != ' ' && line[i] != '\r' && line[i] != '\n')
            return i + 1;
      return -1;
      }

    public void Execute(object sender, EventArgs e)
      {
      textManager.GetActiveView(1, null, out IVsTextView textView);
      textView.GetCaretPos(out int lineOld, out int columnOld);
      dte.ExecuteCommand("Edit.WordNext");
      textView.GetCaretPos(out int lineNew, out int columnNew);

      if (lineNew == lineOld)
        {
        IVsTextLines lines = null;
        textView.GetBuffer(out lines);
        lines.GetLengthOfLine(lineOld, out int lineLength);
        if (lineLength != columnNew)
          return;
        }

      textView.GetTextStream(lineOld, columnOld, lineOld + 1, 0, out string line);             
      int lastPosition = GetLastPosition(line);
      if (lastPosition == -1)
        {
        textView.SetCaretPos(lineOld, 0);              
        if (columnOld == 0)
          dte.ExecuteCommand("Edit.LineEnd");
        else
          dte.ExecuteCommand("Edit.LineStart");
        }
      else 
        textView.SetCaretPos(lineOld, columnOld + lastPosition);
      }
    }
  }







