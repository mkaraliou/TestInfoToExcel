using System.Reflection;

namespace GetterTestInfo_Pvs.Taf
{
    public class MainClass
    {
        public static void Main(string[] args)
        {
            var testProjectName = args[0].Split('\\').Last().Split(".dll").First();

            Assembly sampleAssembly = Assembly.LoadFrom(args[0]);

            var types = sampleAssembly.GetTypes().Where(t => t.FullName.Contains($"{testProjectName}.Tests")).ToList();

            var msethods = types.SelectMany(t => t.GetMethods()).ToList();

            var excelWorker = new ExcelWorker();
            excelWorker.CreateExcelFileWithTestInfos(types);
        }
    }
}
