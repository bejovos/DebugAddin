﻿using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Controls;

namespace DebugAddin.CmdArgs
  {

  public partial class CommandlineArgsToolWindowControl : UserControl, IVsSolutionEvents
    {
    DTE2 dte = (DTE2)Package.GetGlobalService(typeof(DTE));

    List<string> roots = new List<string>{ }; // paths to MatSDK and/or MDCK
    string testSystemRoot = null;
    DataBaseRefresher.DataBase dataBase = null;

    static private void SaveCmdArgs(DataTable dataTable)
      {
      DTE2 dte = (DTE2)Package.GetGlobalService(typeof(DTE));
      var file = new StreamWriter(dte.Solution.FullName + @".debugaddin.cmdargs");
      file.WriteLine("Version 1.2");
      foreach (DataRow row in dataTable.Rows)
        {
        file.WriteLine(row["CommandArguments"]);
        file.WriteLine(row["Filename"]);
        }
      file.Close();
      }

    private void SaveRoots()
      {
      var file = new StreamWriter(dte.Solution.FullName + @".debugaddin.roots");
      file.WriteLine("Version 1.2");
      foreach (string root in roots)
        {
        file.WriteLine(root);
        }
      file.Close();
      }

    private void LoadRoots()
      {
      try
        {
        if (File.Exists(dte.Solution.FullName + @".debugaddin.roots"))
          {
          string[] lines = File.ReadAllLines(dte.Solution.FullName + @".debugaddin.roots");
          if (lines.Length > 0 && lines[0] == "Version 1.2")
            {
            roots = lines.Skip(1).ToList();
            }
          }
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace);
        }

      testSystemRoot = null;
      foreach (string root in roots)
        if (Directory.Exists(root + @"\AlgoTester\Utils\"))
          testSystemRoot = root + @"\AlgoTester\";
      }

    private void LoadSettings(bool quiet)
      {
      // CmdArgs

      DataTable dataTable = new DataTable();
      dataTable.Columns.Add("CommandArguments");
      dataTable.Columns.Add("Filename");
      DTE2 dte = (DTE2)Package.GetGlobalService(typeof(DTE));

      try
        {
        if (File.Exists(dte.Solution.FullName + @".debugaddin.cmdargs"))
          {
          string[] lines = File.ReadAllLines(dte.Solution.FullName + @".debugaddin.cmdargs");
          if (lines.Length > 0 && lines[0] == "Version 1.2")
            {
            foreach (var array in Utils.TupleUp(lines.Skip(1), 2))
              dataTable.Rows.Add(array[0], array[1]);
            }
          }
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace);
        }

      dataTable.RowChanged += (x, y) => SaveCmdArgs(dataTable);
      dataTable.RowDeleted += (x, y) => SaveCmdArgs(dataTable);
      dataGrid.ItemsSource = dataTable.DefaultView;

      // Roots

      LoadRoots();

      // Cases

      try
        {
        dataBase = new DataBaseRefresher.DataBase();
        if (File.Exists(dte.Solution.FullName + @".debugaddin.algotests"))
          dataBase.LoadFromFile(dte.Solution.FullName + @".debugaddin.algotests");
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace);
        }

      if (quiet == false)
        System.Windows.Forms.MessageBox.Show("Loaded!");
      }

    BuildEvents buildEvents;
    CommandEvents commandDebugStartWithoutDebuggingEvents;
    CommandEvents commandDebugStartDebugTargetEvents;
    CommandEvents commandDebugStartEvents;
    public CommandlineArgsToolWindowControl()
      {
      InitializeComponent();

      uint cookie;
      (Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution).AdviseSolutionEvents(this, out cookie);

      buildEvents = dte.Events.BuildEvents;
      buildEvents.OnBuildDone += BuildEvents_OnBuildDone;
      
      var command = dte.Commands.Item("Debug.StartWithoutDebugging");
      commandDebugStartWithoutDebuggingEvents = dte.Events.CommandEvents[command.Guid, command.ID];
      commandDebugStartWithoutDebuggingEvents.BeforeExecute += BeforeDebugStarted;
      command = dte.Commands.Item("Debug.StartDebugTarget");
      commandDebugStartDebugTargetEvents = dte.Events.CommandEvents[command.Guid, command.ID];
      commandDebugStartDebugTargetEvents.BeforeExecute += BeforeDebugStarted;
      command = dte.Commands.Item("Debug.Start");
      commandDebugStartEvents = dte.Events.CommandEvents[command.Guid, command.ID];
      commandDebugStartEvents.BeforeExecute += BeforeDebugStarted;
      }

    private void BeforeDebugStarted(
      string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
      {
      try
        {
        var testCase = GetTestCaseFromRow((dataGrid.SelectedItem as DataRowView).Row);
        if (testCase == null)
          return;

        // generate params file
        CreateParamsFile(testCase);
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace);
        }
      }

    class ParsedRow
      {
      public Project project;
      public string fileName;
      public int lineNumber;
      public string command;
      public string commandArguments;
      public DataBaseRefresher.DataBase.TestCase testCase;
      };

    private void ParseFileNameWithLineNumber(string fileNameWithLineIndex, out string fileName, out int lineNumber)
      {
      var matches = new Regex(@"(?<fileName>.*):(?<lineNumber>\d*)$").Match(fileNameWithLineIndex);
      fileName = matches.Groups["fileName"].Value;
      lineNumber = int.Parse(matches.Groups["lineNumber"].Value);
      }

    private ParsedRow ParseRow(DataRow row)
      {
      if (row == null)
        return null;

      string command;
      string commandArguments = row["CommandArguments"] as string;
      
      if (commandArguments == null)
        return null;

      string[] arguments = commandArguments.Split(' ');
      if (arguments.Length == 0)
        return null;

      DataBaseRefresher.DataBase.TestCase testCase = dataBase?.FindCaseByName(arguments[0]);
      string fileName = testCase?.sourceFile;
      Project project = null;
      int lineNumber;
      
      // Example: M3DTriangleTree_collision_0 --numthreads 1
      if (fileName != null)
        {
        ParseFileNameWithLineNumber(fileName, out fileName, out lineNumber);
        project = dte.Solution.FindProjectItem(fileName)?.ContainingProject;

        if (project == null)
          {
          System.Windows.Forms.MessageBox.Show("Testable project is not found:\n" + fileName);
          }

        string utils = testSystemRoot + @"\Utils\";
        string intermediate = testSystemRoot + @"\Intermediate\";
        string caseName = arguments[0];

        caseName = new Regex(@"[ (){}\[\]+]").Replace(caseName, "_");

        commandArguments = @"--binary ""$(TargetPath)"" --input """ +
          intermediate + caseName + @"_statistics.params"" --statistic """ +
          intermediate + caseName + @"_statistics.stats"" " +
          string.Join(" ", arguments.Skip(1));

        command = utils + @"TKCaseLauncher64.exe";
        }
      else
        {
        ParseFileNameWithLineNumber(row["Filename"] as string, out fileName, out lineNumber);
        // Example: -suite M3DTriangleTree
        project = dte.Solution.FindProjectItem(fileName)?.ContainingProject;
        command = "$(TargetPath)";
        }

      return new ParsedRow
        {
        project = project,
        fileName = fileName,
        lineNumber = lineNumber,
        command = command,
        commandArguments = commandArguments,
        testCase = testCase
        };
      }

    private DataBaseRefresher.DataBase.TestCase GetTestCaseFromRow(DataRow row)
      {
      return ParseRow(row)?.testCase;
      }

    private Project GetProjectFromRow(DataRow row)
      {
      if (row == null)
        return null;
      try
        {
        return ParseRow(row).project;
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace);
        }
      return null;
      }

    private void OpenFileFromRow(DataRow row)
      {
      if (row == null)
        return;
      try
        {
        var parsedRow = ParseRow(row);
        dte.ItemOperations.OpenFile(parsedRow.fileName);
        dte.ExecuteCommand("Edit.GoTo " + parsedRow.lineNumber);
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace);
        }
      }

    private void SetStartupFromRow(DataRow row)
      {
      if (row == null)
        return;
      try
        {
        var parsedRow = ParseRow(row);

        if (parsedRow == null || parsedRow.project == null)
          return;

        dynamic vcPrj = (dynamic)parsedRow.project.Object; // is VCProject
        dynamic vcCfg = vcPrj.ActiveConfiguration; // is VCConfiguration
        dynamic vcDbg = vcCfg.DebugSettings;  // is VCDebugSettings
        vcDbg.Command = parsedRow.command;
        vcDbg.CommandArguments = parsedRow.commandArguments;

        UIHierarchyItem item = Utils.FindUIHierarchyItem(parsedRow.project);
        item.Select(vsUISelectionType.vsUISelectionTypeSelect);

        if (dte.Solution.SolutionBuild.BuildState != vsBuildState.vsBuildStateInProgress)
          dte.Solution.Properties.Item("StartupProject").Value = parsedRow.project.Name;
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace);
        }
      }

    private void BuildEvents_OnBuildDone(vsBuildScope Scope, vsBuildAction Action)
      {
      SetStartupFromRow((dataGrid.SelectedItem as DataRowView)?.Row);
      }

    private void DataGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
      {
      SetStartupFromRow((dataGrid.SelectedItem as DataRowView)?.Row);
      }

    private void Row_DoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
      {
      dataGrid.CommitEdit(DataGridEditingUnit.Row, true);
      SetStartupFromRow(((sender as DataGridRow).Item as DataRowView)?.Row);
      OpenFileFromRow(((sender as DataGridRow).Item as DataRowView)?.Row);
      }

    private void DataGrid_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
      {
      if (e.Key == System.Windows.Input.Key.Enter)
        {
        DataRowView row = dataGrid.SelectedItem as DataRowView;
        if (row.IsEdit == false)
          {
          SetStartupFromRow(row?.Row);
          OpenFileFromRow(row?.Row);
          }
        else
          {
          dataGrid.CommitEdit(DataGridEditingUnit.Row, true);
          }

        e.Handled = true;
        }
      }

    void UpdateRow(DataRow row)
      {
      try
        {
        if (row["Filename"] == null || row["Filename"] as string == null || row["Filename"] as string == "")
          {
          int lineNumber = (dte.ActiveDocument.Selection as TextSelection).ActivePoint.Line;
          row["Filename"] = dte.ActiveDocument.FullName + ":" + lineNumber;
          }
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace);
        }
      }

    private void DataGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
      {
      DataRowView row = dataGrid.SelectedItem as DataRowView;
      UpdateRow(row.Row);
      SetStartupFromRow(row?.Row);
      }

    int IVsSolutionEvents.OnAfterCloseSolution(object pUnkReserved)
      {
      testSystemRoot = null;
      roots.Clear();
      dataBase = null;
      dataGrid.ItemsSource = null;
      return VSConstants.S_OK;
      }

    int IVsSolutionEvents.OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
      {
      LoadSettings(true);

      return VSConstants.S_OK;
      }

    #region IVsSolutionEvents

    int IVsSolutionEvents.OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
      {
      return VSConstants.S_OK;
      }

    int IVsSolutionEvents.OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
      {
      return VSConstants.S_OK;
      }

    int IVsSolutionEvents.OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
      {
      return VSConstants.S_OK;
      }

    int IVsSolutionEvents.OnBeforeCloseSolution(object pUnkReserved)
      {
      return VSConstants.S_OK;
      }

    int IVsSolutionEvents.OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
      {
      return VSConstants.S_OK;
      }

    int IVsSolutionEvents.OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
      {
      return VSConstants.S_OK;
      }

    int IVsSolutionEvents.OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
      {
      return VSConstants.S_OK;
      }

    int IVsSolutionEvents.OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
      {
      return VSConstants.S_OK;
      }

    #endregion

    static bool databaseIsRefreshing = false;

    private async void MenuItem_RecreateDataBase_Click(object sender, System.Windows.RoutedEventArgs e)
      {
      if (databaseIsRefreshing)
        return;
      databaseIsRefreshing = true;
      try
        {
        await System.Threading.Tasks.Task.Run(() => new DataBaseRefresher().RefreshDataBase(roots, dte.Solution.FullName + @".debugaddin.algotests") );
        System.Windows.Forms.MessageBox.Show("Recreated!");
        LoadSettings(false);
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace);
        }
      databaseIsRefreshing = false;
      }

    private void CreateParamsFile(DataBaseRefresher.DataBase.TestCase testCase)
      {
      if (testCase != null)
        {        
        string paramsFile = testSystemRoot + @"\Intermediate\AlgoTesterParams.input";
        Directory.CreateDirectory(Path.GetDirectoryName(paramsFile));
        var file = new StreamWriter(paramsFile);
        file.WriteLine("--cases_root");
        file.WriteLine(testCase.caseFolder);
        file.WriteLine("--settings");
        file.WriteLine(testSystemRoot + @"\settings.config");
        file.WriteLine("--bins_root");
        file.WriteLine(testSystemRoot);
        file.WriteLine("--restrictions");
        file.WriteLine("");
        file.WriteLine("--case");
        file.WriteLine(testCase.caseName);
        file.Close();
        System.Diagnostics.Process pProcess = new System.Diagnostics.Process();
        pProcess.StartInfo.FileName = testSystemRoot + @"\Utils\AlgoTester.exe";
        pProcess.StartInfo.Arguments = "--paramsfile \"" + paramsFile + "\"";
        pProcess.StartInfo.UseShellExecute = false;
        pProcess.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
        pProcess.StartInfo.CreateNoWindow = true;
        pProcess.Start();
        pProcess.WaitForExit();
        }
      }

    private void MenuItem_LoadSettings_Click(object sender, System.Windows.RoutedEventArgs e)
      {
      LoadSettings(false);
      }

    private void MenuItem_EditInputConfig_Click(object sender, System.Windows.RoutedEventArgs e)
      {
      try
        {
        var testCase = GetTestCaseFromRow((dataGrid.SelectedItem as DataRowView).Row);
        if (testCase == null)
          return;

        dte.ItemOperations.OpenFile(testCase.caseFolder + @"\input.config");
        }
      catch (Exception ex)
        {        
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace);
        }
      }

    private void DataGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
      {
      try
        {
        DataRowView row = (DataRowView)e.Row.Item;
        if (!(row.Row["CommandArguments"] is string && (string)row.Row["CommandArguments"] != ""))
          row.Row["CommandArguments"] = dte.Solution.FindProjectItem(dte.ActiveDocument.FullName)?.ContainingProject?.Name;
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace);
        }
      }

    private void MenuItem_AddNewTestRoot_Click(object sender, System.Windows.RoutedEventArgs e)
      {
      try
        {
        using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
          {
          dialog.Description = "Specify path to AlgoTester root (example ...\\MatSDK\\AlgoTester\\)";
          var result = dialog.ShowDialog();
          if (result == System.Windows.Forms.DialogResult.OK || result == System.Windows.Forms.DialogResult.Yes)
            {
            roots.Add(Directory.GetParent(dialog.SelectedPath).FullName);
            SaveRoots();
            LoadRoots();
            }
          }      
        }
      catch (Exception ex)
        {        
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace);
        }
      }
    }
  }