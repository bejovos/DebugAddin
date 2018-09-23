﻿using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using EnvDTE;
using EnvDTE80;
using Task = System.Threading.Tasks.Task;

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

    public void Execute(object sender, EventArgs e)
      {
      TextSelection sel= (TextSelection)dte.ActiveDocument.Selection;
      int columnBefore = sel.CurrentColumn;
      sel.WordRight();
      if (sel.CurrentColumn == 1)
        {
        var editPoint = sel.ActivePoint.CreateEditPoint();
        string text = editPoint.GetText(-2);
        if (text == " \r\n")
          {
          sel.CharLeft();
          sel.WordLeft();
          if (sel.CurrentColumn == columnBefore)
            {
            // rollback our fix if come to the same place 
            sel.MoveToPoint(editPoint);
            }
          }
        }
      }
    }
  }






