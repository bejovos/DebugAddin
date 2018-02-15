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

    int processId = 0;

    string global_expression;
    int global_expression_counter;
    string global_shell_file;
    void Dump(int expression_counter_old)
      {
      if (expression_counter_old != global_expression_counter)
        return;
      Stopwatch sw = new Stopwatch();
      sw.Start();
      
      string constructed_expression = 
        "((int(*)(wchar_t*))Dumper.dll!DumpV)(" +
         "L\"" + global_expression_counter.ToString() + global_expression + "\")";
      Utils.PrintMessage("Dumper", "Constructed expression: " + constructed_expression);
      Utils.PrintMessage("Dumper", dte.Debugger.GetExpression(constructed_expression).Value);
      global_expression = "";
      global_expression_counter = 0;
      try
        {
        foreach (string line in File.ReadAllLines(global_shell_file))
          System.Diagnostics.Process.Start(line);
        }
      catch (Exception) { }

      sw.Stop();
      Utils.PrintMessage("Dumper", "Elapsed time: " + sw.Elapsed.ToString());
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

        string expression;
        if (processId != dte.Debugger.CurrentProcess.ProcessID)
          {
          ulong address = 0;
          var process = (EnvDTE90.Process3)dte.Debugger.CurrentProcess;
          foreach (EnvDTE90.Module module in process.Modules)
            if (module.Name.ToLower() == "kernel32.dll")
              {
              address = module.LoadAddress + 0x1DDD0; // sorry:)
              break;
              }

          expression = "((void*(*)(wchar_t*))"+address+")(L\"Dumper.dll\")";
          Utils.PrintMessage("Dumper", "Constructed expression: " + expression);
          expression = dte.Debugger.GetExpression(expression).Value;
          Utils.PrintMessage("Dumper", expression);
          try
            {
            if (Convert.ToUInt64(expression, 16) == 0)
              throw new Exception("");
            }
          catch (Exception)
            {
            Utils.PrintMessage("Dumper", "Load Library failure!");
            return 0;
            }

          expression = "((int(*)())Dumper.dll!GetDumperVersion)()";
          Utils.PrintMessage("Dumper", "Constructed expression: " + expression);
          expression = dte.Debugger.GetExpression(expression).Value;
          Utils.PrintMessage("Dumper", expression);
          int version;
          if (int.TryParse(expression, out version) == false || version <= 1)
            {
            Utils.PrintMessage("Dumper", "Incompatible version!");
            return 0;
            }
          
          global_shell_file = Path.GetTempFileName();
          expression = "((int(*)(wchar_t*))Dumper.dll!SetShellFile)(L\"" + 
            global_shell_file.Replace("\\","\\\\") + "\")";
          Utils.PrintMessage("Dumper", "Constructed expression: " + expression);
          expression = dte.Debugger.GetExpression(expression).Value;
          Utils.PrintMessage("Dumper", expression);
          if (expression != "0")
            {
            Utils.PrintMessage("Dumper", "Shell file is not set!");
            }

          Utils.PrintMessage("Dumper", "Dumper.dll!Help() for help");

          global_expression = "";
          global_expression_counter = 0;
          processId = process.ProcessID;
          }

        int index = propertyInfo[0].bstrType.IndexOf(" {");
        if (index != -1)
          propertyInfo[0].bstrType = propertyInfo[0].bstrType.Remove(index);
        bool isPointer = IsEndsWith(propertyInfo[0].bstrType, "*") || IsEndsWith(propertyInfo[0].bstrType, "* const");

        string variableName = propertyInfo[0].bstrName;
        if (propertyInfo[0].bstrFullName != propertyInfo[0].bstrName)
          variableName = new Regex(@"^[^$.-]*").Match(propertyInfo[0].bstrFullName).Value 
            + "..." + variableName;
        variableName = variableName.Replace("\"", " ").Replace("'"," ").Replace("\\"," ").Replace("/"," ")
          .Replace(" ", "_");

        expression = (isPointer ? "" : "&") + propertyInfo[0].bstrFullName;
        Utils.PrintMessage("Dumper", "Constructed expression: " + expression);
        expression = dte.Debugger.GetExpression(expression).Value;
        if (expression == "0" || expression == "???")
          {
          Utils.PrintMessage("Dumper", expression);
          Utils.PrintMessage("Dumper", "Incorrect argument is ignored");
          return 0;
          }
        Utils.PrintMessage("Dumper", expression);

        global_expression = global_expression 
          + " " + visualizerId.ToString() 
          + " " + expression
          + " " + variableName;
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
        Utils.PrintMessage("Dumper", ex.Message);
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
      }

    #endregion
    }
  }
