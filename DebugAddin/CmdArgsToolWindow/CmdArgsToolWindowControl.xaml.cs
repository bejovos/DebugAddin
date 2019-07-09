using EnvDTE;
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

namespace DebugAddin.CmdArgsToolWindow
  {

  public partial class CmdArgsToolWindowControl : UserControl, IVsSolutionEvents
    {
    public static CmdArgsToolWindowControl Instance = null;

    public static void Initialize()
      {
      Instance = new CmdArgsToolWindowControl();
      }

    DTE2 dte = (DTE2)Package.GetGlobalService(typeof(DTE));

    List<string> roots = new List<string> { }; // paths to MatSDK and/or MDCK
    string testSystemRoot = null;
    DataBaseRefresher.DataBase dataBase = null;

    static private void SaveCmdArgs(DataTable dataTable)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
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
      ThreadHelper.ThrowIfNotOnUIThread();
      var file = new StreamWriter(dte.Solution.FullName + @".debugaddin.roots");
      file.WriteLine("Version 1.3");
      file.WriteLine(testSystemRoot);

      foreach (string root in roots)
        {
        file.WriteLine(root);
        }
      file.Close();
      }

    private void LoadRoots()
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      roots = null;
      testSystemRoot = null;
      try
        {
        if (File.Exists(dte.Solution.FullName + @".debugaddin.roots"))
          {
          string[] lines = File.ReadAllLines(dte.Solution.FullName + @".debugaddin.roots");
          if (lines.Length > 0 && lines[0] == "Version 1.3")
            {
            testSystemRoot = lines[1];
            roots = lines.Skip(2).ToList();
            }
          }
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace, true);
        }
      }

    private void LoadSettings(bool quiet)
      {
      // CmdArgs
      ThreadHelper.ThrowIfNotOnUIThread();

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
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace, true);
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
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace, true);
        }

      if (quiet == false)
        System.Windows.Forms.MessageBox.Show("Loaded!");
      }

    BuildEvents buildEvents;
    public CmdArgsToolWindowControl()
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      InitializeComponent();

      uint cookie;
      (Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution).AdviseSolutionEvents(this, out cookie);

      buildEvents = dte.Events.BuildEvents;
      buildEvents.OnBuildDone += BuildEvents_OnBuildDone;
      }
   
    class ParsedRow
      {
      public Project project = null;
      public string title;
      public string fileName; // file in solution associated with this row
      public int lineNumber; // line in file
      public string command; // command for debugging 
      public string commandArguments; // arguments
      public DataBaseRefresher.DataBase.TestCase testCase = null; // test case
      };

    private void ParseFileNameWithLineNumber(string fileNameWithLineIndex, out string fileName, out int lineNumber)
      {
      var matches = new Regex(@"(?<fileName>.*):(?<lineNumber>\d*)$").Match(fileNameWithLineIndex);
      fileName = matches.Groups["fileName"].Value;
      lineNumber = int.Parse(matches.Groups["lineNumber"].Value);
      }

    private string GetFirstArgument(ref string commandLine)
      {
      bool escaped = false;
      string result = "";
      int i = 0;

      commandLine = commandLine.Trim();
      for (; i < commandLine.Length; ++i)
        {
        if (!escaped && commandLine[i] == ' ')
          break;

        if (commandLine[i] == '"')
          escaped = ! escaped;
        else
          result += commandLine[i];
        }

      commandLine = commandLine.Substring(i).Trim();

      return result;
      }

    private string ReplaceBackslashesWithSlashes(string str)
      { 
      return str.Replace('\\', '/');
      }

    private ParsedRow ParseRow(DataRow row)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      if (row == null)
        return null;
      string arguments = row["CommandArguments"] as string;
      if (arguments == null)
        return null;

      ParsedRow parsedRow = new ParsedRow();
      parsedRow.title = arguments.Replace("-", "‐"); // replcace &#45 with &#8208

      string restArguments = arguments;
      string firstArgument = GetFirstArgument(ref restArguments);

      parsedRow.testCase = dataBase?.FindCaseByName(firstArgument);

      // Example: M3DTriangleTree_collision_0 --numthreads 1
      if (parsedRow.testCase != null && testSystemRoot != null)
        {
        ParseFileNameWithLineNumber(parsedRow.testCase.sourceFile, out parsedRow.fileName, out parsedRow.lineNumber);
        parsedRow.project = dte.Solution.FindProjectItem(parsedRow.fileName)?.ContainingProject;

        if (parsedRow.project == null)
          {
          System.Windows.Forms.MessageBox.Show("Testable project is not found:\n" + parsedRow.fileName);
          }

        parsedRow.command = testSystemRoot + @"AlgoTester.exe";
        parsedRow.commandArguments = 
          @"--cases_root """ + ReplaceBackslashesWithSlashes(parsedRow.testCase.caseFolder) + @""" " +
          @"--settings """ + ReplaceBackslashesWithSlashes(testSystemRoot) + @"../settings.config"" " +
          @"--bins_root """ + ReplaceBackslashesWithSlashes(((dynamic)parsedRow.project.Object).ActiveConfiguration.OutputDirectory) + @""" " +
          @"--case " + parsedRow.testCase.caseName + @" " + 
          @"--additionaltime 86400000 " +
          restArguments;
        }
      else
        {
        ParseFileNameWithLineNumber(row["Filename"] as string, out parsedRow.fileName, out parsedRow.lineNumber);
        foreach (Project p in Utils.GetAllProjectsInSolution())
          if (p.Name == firstArgument)
            {
            // Example: Project.Tests -suite UserTest
            parsedRow.project = p;
            }
            
        if (parsedRow.project == null)
          {
          // Example: -test UserTest
          parsedRow.project = dte.Solution.FindProjectItem(parsedRow.fileName)?.ContainingProject;
          parsedRow.commandArguments = arguments;
          }
        else
          parsedRow.commandArguments = restArguments;

        parsedRow.command = "$(TargetPath)";
        }

      return parsedRow;
      }

    private void OpenFileFromRow(DataRow row)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
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
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace, true);
        }
      }

    public CmdArgsToolWindow toolwindow;
    private void SetStartupFromRow(ParsedRow parsedRow)
      { 
      ThreadHelper.ThrowIfNotOnUIThread();
      if (parsedRow == null || parsedRow.project == null)
        return;
      try
        {
        dynamic vcPrj = (dynamic)parsedRow.project.Object; // is VCProject
        dynamic vcCfg = vcPrj.ActiveConfiguration; // is VCConfiguration
        dynamic vcDbg = vcCfg.DebugSettings;  // is VCDebugSettings
        vcDbg.Command = parsedRow.command;
        vcDbg.CommandArguments = parsedRow.commandArguments;
        toolwindow.Caption = "Args: " + parsedRow.title;
        UIHierarchyItem item = Utils.FindUIHierarchyItem(parsedRow.project);
        item.Select(vsUISelectionType.vsUISelectionTypeSelect);
        if (dte.Solution.SolutionBuild.BuildState != vsBuildState.vsBuildStateInProgress)
          dte.Solution.Properties.Item("StartupProject").Value = parsedRow.project.Name;
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace, true);
        }        
      }

    private void SetStartupFromRow(DataRow row)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      if (row == null)
        return;
      try
        {
        SetStartupFromRow(ParseRow(row));
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace, true);
        }
      }

    private void BuildEvents_OnBuildDone(vsBuildScope Scope, vsBuildAction Action)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      SetStartupFromRow((dataGrid.SelectedItem as DataRowView)?.Row);
      }

    private void DataGrid_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      SetStartupFromRow((dataGrid.SelectedItem as DataRowView)?.Row);
      }

    private void Row_DoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      dataGrid.CommitEdit(DataGridEditingUnit.Row, true);
      SetStartupFromRow(((sender as DataGridRow).Item as DataRowView)?.Row);
      OpenFileFromRow(((sender as DataGridRow).Item as DataRowView)?.Row);
      }

    private void DataGrid_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
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
      ThreadHelper.ThrowIfNotOnUIThread();
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
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace, true);
        }
      }

    private void DataGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
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
      ThreadHelper.ThrowIfNotOnUIThread();
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
      await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
      if (databaseIsRefreshing)
        return;
      databaseIsRefreshing = true;
      try
        {
        string solutionName = dte.Solution.FullName;
        await System.Threading.Tasks.Task.Run(async () => { 
          await new DataBaseRefresher().RefreshDataBaseAsync(roots, solutionName + @".debugaddin.algotests"); 
          });
        System.Windows.Forms.MessageBox.Show("Recreated!");
        LoadSettings(false);
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace, true);
        }
      databaseIsRefreshing = false;
      }

    private void MenuItem_LoadSettings_Click(object sender, System.Windows.RoutedEventArgs e)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      LoadSettings(false);
      }

    private void MenuItem_EditInputConfig_Click(object sender, System.Windows.RoutedEventArgs e)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      try
        {
        var testCase = ParseRow((dataGrid.SelectedItem as DataRowView).Row)?.testCase;
        if (testCase == null)
          return;

        dte.ItemOperations.OpenFile(testCase.caseFolder + @"\input.config");
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace, true);
        }
      }

    private void DataGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      try
        {
        DataRowView row = (DataRowView)e.Row.Item;
        if (!(row.Row["CommandArguments"] is string && (string)row.Row["CommandArguments"] != ""))
          row.Row["CommandArguments"] = dte.Solution.FindProjectItem(dte.ActiveDocument.FullName)?.ContainingProject?.Name;
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace, true);
        }
      }

    private void MenuItem_AddNewTestRoot_Click(object sender, System.Windows.RoutedEventArgs e)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
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
        Utils.PrintMessage("Exception", ex.Message + "\n" + ex.StackTrace, true);
        }
      }
    }
  }