using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
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

    string rootFolder = "";
    DataBaseRefresher.DataBase dataBase = null;

    static private DataBaseRefresher.DataBase LoadDataBase(string rootFolder)
      {
      try
        {
        if (rootFolder == null || File.Exists(rootFolder + @"\AlgotesterCmd2.tmp") == false)
          return null;
        DataBaseRefresher.DataBase dataBase = new DataBaseRefresher.DataBase();
        dataBase.LoadFromFile(rootFolder + @"\AlgotesterCmd2.tmp");
        return dataBase;
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message);
        }
      return null;
      }

    static private void SaveCmdArgs(DataTable dataTable, string rootFolder)
      {
      DTE2 dte = (DTE2)Package.GetGlobalService(typeof(DTE));
      var file = new StreamWriter(dte.Solution.FullName + @".cmdargs");
      file.WriteLine("Version 1.1");
      file.WriteLine(rootFolder);
      foreach (DataRow row in dataTable.Rows)
        {
        file.WriteLine(row["CommandArguments"]);
        file.WriteLine(row["Filename"]);
        }
      file.Close();
      }

    static private DataTable LoadCmdArgs(out string rootFolder)
      {
      rootFolder = "";
      DataTable dataTable = new DataTable();
      dataTable.Columns.Add("CommandArguments");
      dataTable.Columns.Add("Filename");
      DTE2 dte = (DTE2)Package.GetGlobalService(typeof(DTE));
      string file = dte.Solution.FullName + @".cmdargs";
      if (File.Exists(file))
        {
        string[] lines = File.ReadAllLines(file);
        if (lines.Length > 0 && lines[0] == "Version 1.1")
          {
          rootFolder = lines[1];
          foreach (var array in Utils.TupleUp(lines.Skip(2), 2))
            dataTable.Rows.Add(array[0], array[1]);
          }
        }
      return dataTable;
      }

    private void LoadSettings(bool quiet)
      {
      try
        {
        DataTable dataTable = LoadCmdArgs(out rootFolder);
        dataTable.RowChanged += (x, y) => SaveCmdArgs(dataTable, rootFolder);
        dataTable.RowDeleted += (x, y) => SaveCmdArgs(dataTable, rootFolder);
        dataGrid.ItemsSource = dataTable.DefaultView;
        dataBase = LoadDataBase(rootFolder);
        if (quiet == false)
          System.Windows.Forms.MessageBox.Show("Loaded!");
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message);
        }
      }

    BuildEvents buildEvents;
    public CommandlineArgsToolWindowControl()
      {
      InitializeComponent();

      uint cookie;
      (Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution).AdviseSolutionEvents(this, out cookie);

      buildEvents = dte.Events.BuildEvents;
      buildEvents.OnBuildDone += BuildEvents_OnBuildDone;
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

        string utils = rootFolder + @"\MatSDK\AlgoTester\Utils\";
        string intermediate = rootFolder + @"\MatSDK\AlgoTester\Intermediate\";
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
        fileName = dataBase?.FindSourceFileByOperatorName(arguments[0]);
        // Example: M3DTriangleTree_CollisionDetection
        if (fileName != null)
          {
          ParseFileNameWithLineNumber(fileName, out fileName, out lineNumber);
          project = dte.Solution.FindProjectItem(fileName)?.ContainingProject;
          commandArguments = string.Join(" ", arguments.Skip(1));
          command = "$(TargetPath)";
          }
        else
          {
          ParseFileNameWithLineNumber(row["Filename"] as string, out fileName, out lineNumber);
        
          foreach (Project p in Utils.GetAllProjectsInSolution())
            if (p.Name == arguments[0])
              project = p;
          // Example: -suite M3DTriangleTree
          if (project == null)
            project = dte.Solution.FindProjectItem(fileName)?.ContainingProject;
          // Example: MatSDK.Math.DataQueries.Tests -suite M3DTriangleTree
          else
            commandArguments = string.Join(" ", arguments.Skip(1));
          command = "$(TargetPath)";
          }
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
        Utils.PrintMessage("Exception", ex.Message);
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
        Utils.PrintMessage("Exception", ex.Message);
        }
      }

    private void SetStartupFromRow(DataRow row)
      {
      if (row == null)
        return;
      try
        {
        var parsedRow = ParseRow(row);

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
        Utils.PrintMessage("Exception", ex.Message);
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
        Utils.PrintMessage("Exception", ex.Message);
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
      rootFolder = "";
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
        await System.Threading.Tasks.Task.Run(() => new DataBaseRefresher().RefreshDataBase(rootFolder));
        System.Windows.Forms.MessageBox.Show("Recreated!");
        LoadSettings(false);
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message);
        }
      databaseIsRefreshing = false;
      }

    private void MenuItem_Build_Click(object sender, System.Windows.RoutedEventArgs e)
      {
      try
        {
        Project project = GetProjectFromRow((dataGrid.SelectedItem as DataRowView).Row);
        UIHierarchyItem item = Utils.FindUIHierarchyItem(project);
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
        Utils.PrintMessage("Exception", ex.Message);
        }
      }

    private void MenuItem_GenerateParamsFile_Click(object sender, System.Windows.RoutedEventArgs e)
      {
      try
        {
        var testCase = GetTestCaseFromRow((dataGrid.SelectedItem as DataRowView).Row);
        if (testCase != null)
          {
          string paramsFile = rootFolder + @"\MatSDK\AlgoTester\Intermediate\AlgoTesterParams.input";
          var file = new StreamWriter(paramsFile);
          file.WriteLine("--cases_root");
          file.WriteLine(testCase.caseFolder);
          file.WriteLine("--settings");
          file.WriteLine(rootFolder + @"\MatSDK\AlgoTester\settings.config");
          file.WriteLine("--bins_root");
          file.WriteLine(rootFolder);
          file.WriteLine("--restrictions");
          file.WriteLine("");
          file.WriteLine("--case");
          file.WriteLine(testCase.caseName);
          file.Close();
          System.Diagnostics.Process pProcess = new System.Diagnostics.Process();
          pProcess.StartInfo.FileName = rootFolder + @"\MatSDK\AlgoTester\Utils\AlgoTester.exe";
          pProcess.StartInfo.Arguments = "--paramsfile \"" + paramsFile + "\"";
          pProcess.StartInfo.UseShellExecute = false;
          pProcess.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
          pProcess.StartInfo.CreateNoWindow = true;
          pProcess.Start();
          pProcess.WaitForExit();
          }
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Exception", ex.Message);
        }
      }

    private void MenuItem_LoadSettings_Click(object sender, System.Windows.RoutedEventArgs e)
      {
      LoadSettings(false);
      }
    }
  }