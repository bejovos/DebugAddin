﻿using System;
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
    IVsTextManager2 textManager = (IVsTextManager2)Package.GetGlobalService(typeof(SVsTextManager));

    private void Execute(object sender, EventArgs e)
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
      int whitespaceWordCount = 0;
      if (NextWordCommand.previousCommand == 0b0)
        {
        whitespaceWordCount += 1;
        }
      NextWordCommand.previousCommand = ((NextWordCommand.previousCommand << 1) + 0) & 1;

      if (index > lineLength)
        index = lineLength;

      if (index <= frontWhitespaces)
        {
        if (lineLength == 0)
          dte.ExecuteCommand("Edit.LineEnd");
        else
          {
          index = lineLength;
          textView.SetCaretPos(line, index);
          }
        return;
        }

      TextSpan[] span = new TextSpan[1];
      for (; index > frontWhitespaces && wordsCount != 0; wordsCount -= 1)
        {
        if (buffer[index - 1] == ' ')
          {
          while (index > frontWhitespaces && buffer[index - 1] == ' ')
            index -= 1;
          if (whitespaceWordCount > 0)
            {
            --whitespaceWordCount;
            }
          else 
            continue;
          }

        index -= 1;
        textView.GetWordExtent(line, index, (uint)WORDEXTFLAGS.WORDEXT_PREVIOUS + (uint)WORDEXTFLAGS.WORDEXT_FINDTOKEN, span);
        if (span[0].iStartIndex != span[0].iEndIndex)
          { 
          if (index != span[0].iEndIndex)  
            index = span[0].iStartIndex;
          }
        
        continue;
        }

      textView.SetCaretPos(line, index);
      }
    }
  }
