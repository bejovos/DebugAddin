using System;
using System.ComponentModel.Design;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Debugger.Interop;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using System.IO;
using Microsoft.VisualStudio.Shell.Interop;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace DebugAddin
  {
  [Guid("0BC60053-EB77-46F3-A0B8-F8C41D930D56")]
  public interface IRuntimeDumperService { }

  internal class RuntimeDumperService : IVsCppDebugUIVisualizer, IRuntimeDumperService
    {
    DTE2 dte = (DTE2)Package.GetGlobalService(typeof(DTE));

    bool IsEndsWith(string s, string sub)
      {
      return (s.Substring(s.Length - sub.Length) == sub);
      }

    string ExecuteExpression(string expression, bool useAutoExpandRules = false)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      Utils.PrintMessage("Dumper", "[Dumper] Expression: " + expression);
      string value = dte.Debugger.GetExpression(expression, useAutoExpandRules).Value;
      Utils.PrintMessage("Dumper", "[Dumper] Result: " + value);
      return value;
      }

    struct DebugValue
      {
      public uint visualizerId;
      public string variableName; // human readable name of a variable
      public string expressionResult; // expression result, typically a pointer to variable
      public string typeName; // type of variable
      public string externalDumper; // expression to evaluate external dumper, typically a function
      };

    System.Collections.Generic.List<DebugValue> debugValues = new System.Collections.Generic.List<DebugValue>();

    string global_shell_file;
    int processId = 0;
    string load_library_result = "";

    static string RunLibInf64(string arguments)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      System.Diagnostics.Process process = new System.Diagnostics.Process();
      process.StartInfo.FileName = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\LibInf64.exe";
      process.StartInfo.Arguments = arguments;
      Utils.PrintMessage("Dumper", "[Dumper] Running process: " + process.StartInfo.FileName + " " + process.StartInfo.Arguments);
      process.StartInfo.UseShellExecute = false;
      process.StartInfo.RedirectStandardOutput = true;
      process.StartInfo.CreateNoWindow = true;
      process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
      process.Start();
      string result = process.StandardOutput.ReadToEnd();
      Utils.PrintMessage("Dumper", "[Dumper] Result: " + result);
      process.WaitForExit();
      return result;
      }

    string GetBaseAddress(string module_name)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      var process = (EnvDTE90.Process3)dte.Debugger.CurrentProcess;
      foreach (EnvDTE90.Module module in process.Modules)
        if (module.Name.ToLower() == module_name)
          return module.LoadAddress.ToString();
      throw new Exception("Target library is not loaded!");
      }

    static ulong? Parse(string value, int fromBase = 10)
      {
      UInt64 result;
      try
        {
        result = Convert.ToUInt64(value, fromBase);
        }
      catch (Exception)
        {
        return null;
        }
      return result;
      }

    void Dump(int expression_counter_old)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      if (debugValues.Count != expression_counter_old)
        return;
      Stopwatch sw = new Stopwatch();
      sw.Start();
      try
        {
        if (processId != dte.Debugger.CurrentProcess.ProcessID)
          {
          dte.StatusBar.Text = "[Dumper] Loading dumpers...";

          load_library_result = ExecuteExpression(
            "((void*(*)(wchar_t*))(" +
            RunLibInf64("kernel32.dll LoadLibraryW") + "+" + GetBaseAddress("kernel32.dll") +
            @"))(L""Dumper.dll"")");
          if (Parse(load_library_result, 16).GetValueOrDefault(0) == 0)
            throw new Exception("Load Library failure!");

          global_shell_file = ExecuteExpression("((wchar_t*(*)())Dumper.dll!LoadDumpers)()", true);
          global_shell_file = Regex.Match(global_shell_file, @"L""(.*)""$").Groups[1].Value;
          global_shell_file = Regex.Replace(global_shell_file, @"\\(\\|""|')", "$1");
          Utils.PrintMessage("Dumper", "[Dumper] Shell file: " + global_shell_file);
          processId = ((EnvDTE90.Process3)dte.Debugger.CurrentProcess).ProcessID;
          }

        dte.StatusBar.Text = "[Dumper] Dumping...";

        string constructedExpression = "";
        string result = "";
        foreach (var debugValue in debugValues)
          {
          if (debugValue.visualizerId > 1000)
            {
            result = ExecuteExpression(debugValue.externalDumper);
            if (Parse(result, 16).GetValueOrDefault(0) == 0)
              throw new Exception(@"""" + debugValue.externalDumper + @""" is not found!");
            }
          constructedExpression = constructedExpression + 
              " " + debugValue.visualizerId.ToString() +
              " " + debugValue.expressionResult +
              " " + debugValue.variableName +
              " " + debugValue.typeName; 
          
          if (debugValue.visualizerId > 1000)
            constructedExpression = constructedExpression + " " + result;
          }
          
        result = ExecuteExpression(@"((int(*)(wchar_t*))Dumper.dll!Dump)(L""" + 
          debugValues.Count.ToString() + constructedExpression + @""")");

        if (result != "0")
          {
          ExecuteExpression("((void(*)())Dumper.dll!UnloadDumpers)()");
          result = ExecuteExpression(
            "((int(*)(void*))(" +
            RunLibInf64("kernel32.dll FreeLibrary") + "+" + GetBaseAddress("kernel32.dll") +
            "))(" + load_library_result + ")");
          if (Parse(result).GetValueOrDefault(0) == 0)
            throw new Exception("Free Library failure!");
          processId = 0;
          throw new Exception("Dumping error!");
          }

        foreach (string line in File.ReadAllLines(global_shell_file))
          {
          Utils.PrintMessage("Debug", "[Dumper] Saved to: " + line);
          System.Diagnostics.Process.Start(line);
          }

        dte.StatusBar.Text = "[Dumper] Dumped successfully";
        }
      catch (Exception ex)
        {
        dte.StatusBar.Text = "[Dumper] " + ex.Message;
        dte.StatusBar.Highlight(true);
        Utils.PrintMessage("Debug", "[Dumper] [ERROR] " + ex.Message, true);
        Utils.PrintMessage("Dumper", "[Dumper] [ERROR] " + ex.Message + "\n" + ex.StackTrace);
        }
      sw.Stop();
      Utils.PrintMessage("Dumper", "[Dumper] Elapsed time: " + sw.Elapsed.ToString());
      debugValues.Clear();
      }

    public int DisplayValue(uint ownerHwnd, uint visualizerId, IDebugProperty3 debugProperty)
      {
      ThreadHelper.ThrowIfNotOnUIThread();
      try
        {
        DEBUG_PROPERTY_INFO[] propertyInfo = new DEBUG_PROPERTY_INFO[1];
        debugProperty.GetPropertyInfo(
          enum_DEBUGPROP_INFO_FLAGS.DEBUGPROP_INFO_STANDARD |
          enum_DEBUGPROP_INFO_FLAGS.DEBUGPROP_INFO_FULLNAME |
          enum_DEBUGPROP_INFO_FLAGS.DEBUGPROP_INFO_PROP |
          enum_DEBUGPROP_INFO_FLAGS.DEBUGPROP_INFO_VALUE_RAW,
          10 /* Radix */,
          10000 /* Eval Timeout */,
          new IDebugReference2[] { },
          0,
          propertyInfo);

        DebugValue debugValue = new DebugValue();
        debugValue.visualizerId = visualizerId;

        int index = propertyInfo[0].bstrType.IndexOf(" {");
        if (index != -1)
          propertyInfo[0].bstrType = propertyInfo[0].bstrType.Remove(index);
        bool isPointer = IsEndsWith(propertyInfo[0].bstrType, "*") || IsEndsWith(propertyInfo[0].bstrType, "* const");
        bool isReference = IsEndsWith(propertyInfo[0].bstrType, "&");

        string variableName = propertyInfo[0].bstrName;
        if (propertyInfo[0].bstrFullName != propertyInfo[0].bstrName)
          variableName = Regex.Match(propertyInfo[0].bstrFullName, @"^[^$.-]*").Value + "..." + variableName;
        debugValue.variableName = Regex.Replace(variableName, @"[""'\/ ]", "_");

        debugValue.expressionResult = ExecuteExpression((isPointer ? "" : "&") + propertyInfo[0].bstrFullName);
        if (Parse(debugValue.expressionResult, 16).GetValueOrDefault(0) == 0)
          throw new Exception("Incorrect argument!");

        string typeName = null;
        if (visualizerId > 1000)
          {
          debugProperty.GetCustomViewerCount(out uint viewersCount);
          DEBUG_CUSTOM_VIEWER[] viewers = new DEBUG_CUSTOM_VIEWER[viewersCount];
          debugProperty.GetCustomViewerList(0, viewersCount, viewers, out uint _);

          for (uint i=0; i != viewersCount; ++i)
            if (System.Guid.Parse(viewers[i].bstrMetric) == typeof(IRuntimeDumperService).GUID
              && viewers[i].dwID == visualizerId)
              {
              debugValue.externalDumper = viewers[i].bstrMenuName;
              typeName = viewers[i].bstrDescription;
              break;
              }
          if (debugValue.externalDumper is null)
            throw new Exception("External dumper is not found!");
          }
        else if (isReference)
          {
          int length = propertyInfo[0].bstrType.Length;
          typeName = propertyInfo[0].bstrType.Substring(0, length - 1) + " *";
          }
        else if (isPointer)
          typeName = propertyInfo[0].bstrType;
        else
          typeName = propertyInfo[0].bstrType + " *";
        debugValue.typeName = @"\""" + Regex.Replace(typeName, @"[\w.]+!", "") + @"\""";

        debugValues.Add(debugValue);
        int count = debugValues.Count;
        _ = System.Threading.Tasks.Task.Delay(1000).ContinueWith(t =>
            {
              Dump(count);
            }, System.Threading.Tasks.TaskScheduler.FromCurrentSynchronizationContext());
        }
      catch (Exception ex)
        {
        dte.StatusBar.Text = "[Dumper] " + ex.Message;
        dte.StatusBar.Highlight(true);
        Utils.PrintMessage("Debug", "[Dumper] [ERROR] " + ex.Message, true);
        Utils.PrintMessage("Dumper", "[Dumper] [ERROR] " + ex.Message + "\n" + ex.StackTrace);
        }

      return 0;
      }

    public delegate void DumpDelegate(int x);
    }

  [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
  [InstalledProductRegistration("#1110", "#1112", "1.0", IconResourceID = 1400)] // Info on this package for Help/About
  [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
  [Guid("AE816FBD-5FD2-4D75-AF20-0729F3239467")]
  [ProvideService(typeof(IRuntimeDumperService), ServiceName = "RuntimeDumperService")]
  public sealed class RuntimeDumperPackage : AsyncPackage
    {
    public RuntimeDumperPackage()
      {
      base.Initialize();
      IServiceContainer serviceContainer = this;
      if (serviceContainer != null)
        serviceContainer.AddService(typeof(IRuntimeDumperService), new RuntimeDumperService(), true);
      }

    #region Package Members

    /// <summary>
    /// Initialization of the package; this method is called right after the package is sited, so this is the place
    /// where you can put all the initialization code that rely on services provided by VisualStudio.
    /// </summary>
    protected override async System.Threading.Tasks.Task InitializeAsync(System.Threading.CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
      {
      await base.InitializeAsync(cancellationToken, progress);
      }

    #endregion
    }
  }
