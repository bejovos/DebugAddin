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
    public static int previousCommand = 0;

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
    IVsTextManager2 textManager = (IVsTextManager2)Package.GetGlobalService(typeof(SVsTextManager));

    public void Execute(object sender, EventArgs e)
      {
      textManager.GetActiveView2(1, null, (uint)_VIEWFRAMETYPE.vftAny, out IVsTextView textView);
      textView.GetCaretPos(out int line, out int index);
      
      IVsTextLines textLines;
      textView.GetBuffer(out textLines);
      string buffer;
      int lineLength;
      textLines.GetLengthOfLine(line, out lineLength);
      textLines.GetLineText(line, 0, line, lineLength, out buffer);
      while (lineLength > 0 && buffer[lineLength - 1] == ' ')
        lineLength -= 1;
      int frontWhitespaces = 0;
      while (frontWhitespaces < lineLength && buffer[frontWhitespaces] == ' ')
        frontWhitespaces += 1;

      int wordsCount = 1;
      bool includeWhitespacesInWord = false;
      if (previousCommand == 0b1)
        {
        //wordsCount += 1; 
        includeWhitespacesInWord = true;
        }
      previousCommand = ((previousCommand << 1) + 1) & 1;
      if (index >= lineLength)
        {
        if (lineLength == 0)
          dte.ExecuteCommand("Edit.LineEnd");
        else 
          {
          index = frontWhitespaces;
          textView.SetCaretPos(line, index);
          }
        return;
        }
      
      TextSpan[] span = new TextSpan[1];
      for (; index < lineLength && wordsCount != 0; wordsCount -= 1)
        {
        if (buffer[index] == ' ')
          {
          while (index < lineLength && buffer[index] == ' ')
            index += 1;
          continue;  
          }

        textView.GetWordExtent(line, index, (uint)WORDEXTFLAGS.WORDEXT_NEXT + (uint)WORDEXTFLAGS.WORDEXT_FINDTOKEN, span);
        if (span[0].iStartIndex != span[0].iEndIndex)
          { 
          if (index == span[0].iEndIndex)  
            index += 1;
          else
            index = span[0].iEndIndex;
          }
        else
          {
          index += 1;
          }

        if (includeWhitespacesInWord)
          {
          while (index < lineLength && buffer[index] == ' ')
            index += 1;
          }
        continue;
        }
      textView.SetCaretPos(line, index);
      }
    }
  }







