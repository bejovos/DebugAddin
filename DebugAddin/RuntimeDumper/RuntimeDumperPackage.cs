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
using DebugAddin.CmdArgs;

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

    string ExecuteExpression(string expression)
      {
      Utils.PrintMessage("Debug", "[Dumper] Expression: " + expression);
      string value = dte.Debugger.GetExpression(expression).Value;
      Utils.PrintMessage("Debug", "[Dumper] Result: " + value);
      return value;
      }

    string global_expression;
    int global_expression_counter;
    string global_shell_file;
    int processId = 0;
    string load_library_result = "";

    string RunLibInf64(string arguments)
      {
      System.Diagnostics.Process process = new System.Diagnostics.Process();
      process.StartInfo.FileName = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\LibInf64.exe";
      process.StartInfo.Arguments = arguments;
      Utils.PrintMessage("Debug", "[Dumper] Running process: " + process.StartInfo.FileName + " " + process.StartInfo.Arguments);
      process.StartInfo.UseShellExecute = false;
      process.StartInfo.RedirectStandardOutput = true;
      process.StartInfo.CreateNoWindow = true;
      process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
      process.Start();
      string result = process.StandardOutput.ReadToEnd();
      Utils.PrintMessage("Debug", "[Dumper] Result: " + result);
      process.WaitForExit();
      return result;
      }

    string GetBaseAddress(string module_name)
      {
      var process = (EnvDTE90.Process3)dte.Debugger.CurrentProcess;
      foreach (EnvDTE90.Module module in process.Modules)
        if (module.Name.ToLower() == module_name)
          return module.LoadAddress.ToString();
      throw new Exception("Target library is not loaded!");
      }

    void Dump(int expression_counter_old)
      {
      Stopwatch sw = new Stopwatch();
      sw.Start();
      try
        {
        if (expression_counter_old != global_expression_counter)
          return;

        if (processId != dte.Debugger.CurrentProcess.ProcessID)
          {
          load_library_result = ExecuteExpression(
            "((void*(*)(wchar_t*))(" + 
            RunLibInf64("kernel32.dll LoadLibraryW") + "+" + GetBaseAddress("kernel32.dll") + 
            "))(L\"Dumper.dll\")");
          if (Convert.ToUInt64(load_library_result, 16) == 0)
            throw new Exception("Load Library failure!");

          global_shell_file = Path.GetTempFileName();
          string expression = ExecuteExpression("((int(*)(wchar_t*))Dumper.dll!SetShellFile)(L\"" + global_shell_file.Replace("\\", "\\\\") + "\")");
          if (int.Parse(expression) != 0)
            throw new Exception("Shell file is not set!");

          processId = ((EnvDTE90.Process3)dte.Debugger.CurrentProcess).ProcessID;
          }

        string result = ExecuteExpression("((int(*)(wchar_t*))Dumper.dll!DumpV)(L\"" + global_expression_counter.ToString() + global_expression + "\")");

        if (result != "0")
          {
          string expression = ExecuteExpression(
            "((int(*)(void*))(" + 
            RunLibInf64("kernel32.dll FreeLibrary") + "+" + GetBaseAddress("kernel32.dll") + 
            "))(" + load_library_result + ")");
          if (Convert.ToUInt64(expression) == 0)
            throw new Exception("Free Library failure!");
          processId = 0;
          Utils.PrintMessage("Debug", "[Dumper] [ERROR] Dumping error!", true);
          }

        foreach (string line in File.ReadAllLines(global_shell_file))
          System.Diagnostics.Process.Start(line);
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Debug", "[Dumper] " + ex.Message + "\n" + ex.StackTrace, true);
        }
      sw.Stop();
      Utils.PrintMessage("Debug", "[Dumper] Elapsed time: " + sw.Elapsed.ToString());
      global_expression = "";
      global_expression_counter = 0;
      }

    public int DisplayValue(uint ownerHwnd, uint visualizerId, IDebugProperty3 debugProperty)
      {
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

        int index = propertyInfo[0].bstrType.IndexOf(" {");
        if (index != -1)
          propertyInfo[0].bstrType = propertyInfo[0].bstrType.Remove(index);
        bool isPointer = IsEndsWith(propertyInfo[0].bstrType, "*") || IsEndsWith(propertyInfo[0].bstrType, "* const");
        bool isReference = IsEndsWith(propertyInfo[0].bstrType, "&");

        string variableName = propertyInfo[0].bstrName;
        if (propertyInfo[0].bstrFullName != propertyInfo[0].bstrName)
          variableName = new Regex(@"^[^$.-]*").Match(propertyInfo[0].bstrFullName).Value + "..." + variableName;
        variableName = new Regex("[\"'\\/ ]").Replace(variableName, "_");

        string expression = ExecuteExpression((isPointer ? "" : "&") + propertyInfo[0].bstrFullName);
        
        string typeName;
        if (isReference)
          {
          int length = propertyInfo[0].bstrType.Length;
          typeName = propertyInfo[0].bstrType.Substring(0, length - 1) + " *";
          }
        else if (isPointer)
          typeName = propertyInfo[0].bstrType;
        else
          typeName = propertyInfo[0].bstrType + " *";

        typeName = "\\\"" + new Regex(@"[\w.]+!").Replace(typeName, "") + "\\\"";

        if (Convert.ToUInt64(expression, 16) == 0)
          throw new Exception("Incorrect argument!");

        global_expression = global_expression
          + " " + visualizerId.ToString()
          + " " + expression
          + " " + variableName
          + " " + typeName;
        global_expression_counter += 1;

        int expression_counter_value = global_expression_counter;
        var disp = System.Windows.Threading.Dispatcher.CurrentDispatcher;
        System.Threading.Tasks.Task.Delay(1000).ContinueWith(t =>
          {
            disp.BeginInvoke((DumpDelegate)((x) => Dump(x)), expression_counter_value);
          });
        }
      catch (Exception ex)
        {
        Utils.PrintMessage("Debug", "[Dumper] " + ex.Message + "\n" + ex.StackTrace, true);
        }

      return 0;
      }
    public delegate void DumpDelegate(int x);
    }

  [PackageRegistration(UseManagedResourcesOnly = true)]
  [InstalledProductRegistration("#1110", "#1112", "1.0", IconResourceID = 1400)] // Info on this package for Help/About
  [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]
  [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
  [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
  [Guid("AE816FBD-5FD2-4D75-AF20-0729F3239467")]
  [ProvideService(typeof(IRuntimeDumperService), ServiceName = "RuntimeDumperService")]
  public sealed class RuntimeDumperPackage : Package
    {
    public RuntimeDumperPackage()
      {
      base.Initialize();
      IServiceContainer serviceContainer = (IServiceContainer)this;
      if (serviceContainer != null)
        serviceContainer.AddService(typeof(IRuntimeDumperService), new RuntimeDumperService(), true);
      }

    #region Package Members

    /// <summary>
    /// Initialization of the package; this method is called right after the package is sited, so this is the place
    /// where you can put all the initialization code that rely on services provided by VisualStudio.
    /// </summary>
    protected override void Initialize()
      {
      base.Initialize();
      CommandlineArgsToolWindowControl.Initialize();
      }

    #endregion
    }
  }
