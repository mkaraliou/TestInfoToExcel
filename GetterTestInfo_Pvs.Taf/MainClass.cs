using System.Reflection;

namespace GetterTestInfo_Pvs.Taf
{
    public class MainClass
    {
        public static void Main(string[] args)
        {
            //Assembly sampleAssembly = Assembly.LoadFrom(@"C:\Users\Mikalai_Karaliou\source\repos\Pvs.Taf.Core\Pvs.Taf.V2Smoke\bin\Debug\net6.0\Pvs.Taf.V2Smoke.dll");
            Assembly sampleAssembly = Assembly.LoadFrom(args[0]);

            var types = sampleAssembly.GetTypes().Where(t => t.FullName.Contains("Pvs.Taf.V2Smoke.Tests")).ToList();

            var msethods = types.SelectMany(t => t.GetMethods()).ToList();

            var excelWorker = new ExcelWorker();
            excelWorker.CreateExcelFileWithTestInfos(types);
        }
    }
}
