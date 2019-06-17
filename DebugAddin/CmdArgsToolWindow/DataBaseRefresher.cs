using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;
using Microsoft.VisualStudio.Shell;

namespace DebugAddin.CmdArgsToolWindow
  {
  class DataBaseRefresher
    {
    public class DataBase
      {
      public void Add(string caseName, string caseFolder,
        string operatorName, string binaryName, string sourceFile = "")
        {
        TestCase testCase = new TestCase();
        testCase.caseName = caseName;
        testCase.caseFolder = caseFolder;
        testCase.operatorName = operatorName;
        testCase.binaryName = binaryName;
        testCase.sourceFile = sourceFile;
        testCases.Add(caseName, testCase);

        if (operatorToSourceFile.ContainsKey(operatorName) == false)
          operatorToSourceFile.Add(operatorName, sourceFile);
        }
      public void FillSourceFileInfo(
        Dictionary<Tuple<string, string>, string> operatorNameBinaryNameToSourceFile)
        {
        Dictionary<string, List<string>> grouppedCases = testCases.
          GroupBy(r => r.Value.operatorName + "@" + r.Value.binaryName).
          ToDictionary(t => t.Key, t => t.Select(r => r.Key).ToList());

        foreach (var entry in operatorNameBinaryNameToSourceFile)
          {
          string key = entry.Key.Item1 + "@" + entry.Key.Item2;
          if (grouppedCases.ContainsKey(key) == false)
            continue;
          foreach (var entry2 in grouppedCases[key])
            {
            testCases[entry2].sourceFile = entry.Value;
            }
          }
        }
      public void WriteToFile(string fileName)
        {
        var file = new StreamWriter(fileName);
        file.WriteLine("Version 1.2");
        foreach (var testCase in testCases)
          {
          file.WriteLine(testCase.Value.caseName);
          file.WriteLine(testCase.Value.caseFolder);
          file.WriteLine(testCase.Value.operatorName);
          file.WriteLine(testCase.Value.binaryName);
          file.WriteLine(testCase.Value.sourceFile);
          }
        file.Close();
        }
      public void LoadFromFile(string fileName)
        {
        string[] lines = File.ReadAllLines(fileName);
        if (lines.Length > 0 && lines[0] == "Version 1.2")
          {
          foreach (var array in Utils.TupleUp(lines.Skip(1), 5))
            Add(array[0], array[1], array[2], array[3], array[4]);
          }
        }
      public TestCase FindCaseByName(string caseName)
        {
        if (testCases.ContainsKey(caseName) == false)
          return null;
        return testCases[caseName];
        }

      public string FindSourceFileByOperatorName(string operatorName)
        {
        if (operatorToSourceFile.ContainsKey(operatorName) == false)
          return null;
        return operatorToSourceFile[operatorName];
        }

      Dictionary<string, TestCase> testCases = new Dictionary<string, TestCase>(); // key - caseName
      Dictionary<string, string> operatorToSourceFile = new Dictionary<string, string>();

      public class TestCase
        {
        public string caseName; // with incorrect symbols
        public string caseFolder; // full path to concrete folder
        public string operatorName;
        public string binaryName; // with "Testable.dll"
        public string sourceFile; // full path with ":lineNumber"
        };
      };

    private string RemoveQuotes(string input)
      {
      return input.Substring(1, input.Length - 2);
      }

    private int FindOperateLineNumber(string[] lines)
      {
      Regex regex = new Regex(@"::Operate\(\)", RegexOptions.None);
      int operateLineNumber = 0;
      for (int lineNumber = 0; lineNumber < lines.Length; ++lineNumber)
        {
        for (Match match = regex.Match(lines[lineNumber]); match.Success; )
          {
          operateLineNumber = lineNumber;
          break;
          }
        }
      return operateLineNumber;
      }

    private void ProcessTestableFile(string binaryName, string sourceFile)
      {
      string[] lines = File.ReadAllLines(sourceFile);
      int operateLineNumber = FindOperateLineNumber(lines);
      Regex regex = new Regex(@"\""[^\""]*\""", RegexOptions.None);
      var operatorNameBinaryNameToSourceFile = new Dictionary<Tuple<string, string>, string>();
      for (int lineNumber = 0; lineNumber < lines.Length; ++lineNumber)
        {
        for (Match match = regex.Match(lines[lineNumber]); match.Success; match = match.NextMatch())
          {
          string maybeOperatorName = RemoveQuotes(match.Value);
          var tuple = new Tuple<string, string>(maybeOperatorName, binaryName);
          if (operatorNameBinaryNameToSourceFile.ContainsKey(tuple) == false)
            operatorNameBinaryNameToSourceFile.Add(tuple, sourceFile + ":" + 
              ((operateLineNumber == 0 ? lineNumber : operateLineNumber) + 1));
          }
        }
      db.FillSourceFileInfo(operatorNameBinaryNameToSourceFile);
      }

    private async System.Threading.Tasks.Task ProcessTestCaseAsync(string fileName)
      {
      Regex regex = new Regex(@"\""[^\""]*\""", RegexOptions.None);
      string caseName = null;
      string operatorName = null;
      string binaryName = null;
      string caseFolder = Path.GetDirectoryName(fileName);

      Match match = regex.Match(File.ReadAllText(caseFolder + @"\case.config"));
      for (; match.Success; match = match.NextMatch())
        if (match.Value == @"""name""")
          {
          match = match.NextMatch();
          caseName = RemoveQuotes(match.Value);
          }
        else if (match.Value == @"""binary""")
          {
          match = match.NextMatch();
          binaryName = RemoveQuotes(match.Value);
          }

      match = regex.Match(File.ReadAllText(caseFolder + @"\runtime.config"));
      for (; match.Success; match = match.NextMatch())
        if (match.Value == @"""Operator""")
          {
          match = match.NextMatch();
          operatorName = RemoveQuotes(match.Value);
          }

      if (operatorName == null || caseName == null || binaryName == null)
        return;

      try
        {
        db.Add(caseName, caseFolder, operatorName, binaryName);
        }
      catch (Exception ex)
        {
        var testCase = db.FindCaseByName(caseName);
        await Utils.PrintMessageAsync("Exception", caseName + " " + caseFolder + " " + operatorName + " " + binaryName);
        await Utils.PrintMessageAsync("Exception", testCase.caseName + " " + testCase.caseFolder + " " + testCase.operatorName + " " + testCase.binaryName);
        await Utils.PrintMessageAsync("Exception", ex.Message);
        }
      }

    private DataBase db = new DataBase();

    public async System.Threading.Tasks.Task RefreshDataBaseAsync(List<string> roots, string outputFile)
      {
      List<string> directories = new List<string>{ };
      foreach (string root in roots)
        {
        string[] testCases = Directory.GetFiles(root + @"\AlgoTester\Cases\", "case.config", SearchOption.AllDirectories);
        foreach (string testCase in testCases)
          await ProcessTestCaseAsync(testCase);
        directories.AddRange(Directory.GetDirectories(root + @"\Libraries\", "testable", SearchOption.AllDirectories));
        }

      foreach (string directory in directories)
        {
        string binary = Directory.GetParent(directory).Name + @".TestableOperators.dll";
        foreach (string sourceFile in Directory.GetFiles(directory + @"\", @"*", SearchOption.AllDirectories))
          ProcessTestableFile(binary, sourceFile);
        }
      db.WriteToFile(outputFile);
      }
    }
  }
