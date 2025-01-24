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

      Console.WriteLine(process.ToString());

      string processJson = process.ToJson();

      Console.WriteLine(processJson);

    }
  }
}
