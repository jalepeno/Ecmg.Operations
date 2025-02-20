using Documents.Utilities;
using Operations;
using System.Reflection;

namespace OperationTests
{
  [TestClass]
  public sealed class Test1
  {
    [TestMethod]
    public void OpenProcessFromXML()
    {
      Assembly assembly = Assembly.GetExecutingAssembly();
      string processXml = string.Empty;
      using (Stream processConfigStream = assembly.GetManifestResourceStream("OperationTests.Specimens.ProcessSample.xml"))
      {
        using (StreamReader reader = new StreamReader(processConfigStream))
        {
          processXml = reader.ReadToEnd();
        }
      }

      Process process = Process.FromXml(processXml);


      process.RunAfterComplete = new ImportOperation();

      Console.WriteLine(process.ToString());

      string processJson = process.ToJson();

      Console.WriteLine(processJson);

    }

    [TestMethod]
    public void OpenProcessFromJson()
    {
      string jsonPath = "G:\\Deloitte\\EDD\\Test Documents\\newProcess.json";
      string processJson = Helper.ReadAllTextFromFile(jsonPath);
      Process process = Process.CreateFromJson(processJson);
      string readProcessJson = process.ToJson();

    }

  }
}
