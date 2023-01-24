using ClosedXML.Excel;
using NUnit.Framework.Interfaces;
using System.Collections;
using System.Reflection;

namespace GetterTestInfo_Pvs.Taf
{
    public class ExcelWorker
    {
        private  List<string> columns = new List<string> { "TestCaseId", "Class", "Test", "Category", "Priority", "Property", "Description" };

        public void CreateExcelFileWithTestInfos(List<Type> types)
        {
            var methods = types.SelectMany(t => t.GetMethods()).ToList();

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sheet1");

            AddTitleColumns(worksheet);

            var testMethods = methods.Where(m => m.CustomAttributes.Any(a => a.AttributeType.Name == "TestAttribute")).ToList();
            for (int i = 0; i < testMethods.Count; i++)
            {
                FillLineForMethod(worksheet, testMethods[i], i + 2);
            }

            AddDataDrivenTests(types, worksheet, testMethods.Count + 2);

            AddSorting(worksheet);

            //ширина столбца по содержимому
            worksheet.Columns().AdjustToContents(); 
            worksheet.Columns().Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

            workbook.SaveAs($"{DateTime.Now.ToString("yyyy-MM-dd HH.mm.ss")} Smoke.xlsx");
        }

        private void AddTitleColumns(IXLWorksheet worksheet)
        {
            for (int i = 0; i < columns.Count; i++)
            {
                worksheet.Cell(GetLetterByNumber(i) + 1).Value = columns[i];
            }
        }

        private void FillLineForMethod(IXLWorksheet worksheet, MethodInfo method, int lineNumber)
        {
            FillCellClass(worksheet, lineNumber, method.ReflectedType.Name);
            FillCellTestName(worksheet, lineNumber, method.Name);
            FillCellValuesFromConstructor(worksheet, method, "TestCaseId", lineNumber);
            FillCellCategory(worksheet, method, lineNumber);
            FillCellValuesFromConstructor(worksheet, method, "Priority", lineNumber);
            FillCellValuesFromConstructor(worksheet, method, "Description", lineNumber);
            FillCellProperty(worksheet, method, lineNumber);
        }

        private void AddDataDrivenTests(List<Type> types, IXLWorksheet worksheet, int lineNumber)
        {
            var methods = types.SelectMany(t => t.GetMethods()).ToList();
            var dataDrivenTests = methods.Where(t => t.CustomAttributes.Any(a => a.AttributeType.Name == "TestCaseSourceAttribute")).ToList();

            foreach (var method in dataDrivenTests)
            {
                var className = method.ReflectedType.Name;
                var methodName = method.Name;

                var testDataMethodName = method.CustomAttributes.First(a => a.AttributeType.Name == "TestCaseSourceAttribute").ConstructorArguments[0].Value;

                var testDataMethod = types.First(t => t.Name == className).GetMethod(
                    testDataMethodName.ToString(),
                    BindingFlags.Static | BindingFlags.NonPublic);

                var testDataValues = (IEnumerable)testDataMethod.Invoke(null, null);

                AddDataDrivenTest(worksheet, lineNumber, method, testDataValues);
            }
        }

        private void AddDataDrivenTest(IXLWorksheet worksheet, int lineNumber, MethodInfo method, IEnumerable testDataValues)
        {
            foreach (var item in testDataValues)
            {
                FillCellCategory(worksheet, method, lineNumber);
                FillCellClass(worksheet, lineNumber, method.ReflectedType.Name);

                var testCaseData = item as ITestCaseData;

                worksheet.Cell(GetLetterByNumber(columns.IndexOf("Test")) + lineNumber).Value = $"{method.Name} : {testCaseData.TestName}";
                worksheet.Cell(GetLetterByNumber(columns.IndexOf("TestCaseId")) + lineNumber).Value = testCaseData.Properties.Get("TestCaseId").ToString();
                worksheet.Cell(GetLetterByNumber(columns.IndexOf("Priority")) + lineNumber).Value = testCaseData.Properties.Get("Priority").ToString();
                worksheet.Cell(GetLetterByNumber(columns.IndexOf("Description")) + lineNumber).Value = testCaseData.Properties.Get("Description").ToString();

                lineNumber++;
            }

        }

        private void FillCellClass(IXLWorksheet worksheet, int lineNumber, string value)
        {
            var cellLetter = GetLetterByNumber(columns.IndexOf("Class"));
            worksheet.Cell(cellLetter + lineNumber).Value = value;
        }

        private void FillCellTestName(IXLWorksheet worksheet, int lineNumber, string testName)
        {
            var cellLetter = GetLetterByNumber(columns.IndexOf("Test"));
            worksheet.Cell(cellLetter + lineNumber).Value = testName;
        }

        private void FillCellValuesFromConstructor(IXLWorksheet worksheet, MethodInfo method, string columnName, int lineNumber)
        {
            var customAttribute = method.CustomAttributes.FirstOrDefault(a => a.AttributeType.Name.Contains(columnName));
            var cellLetter = GetLetterByNumber(columns.IndexOf(columnName));

            if (customAttribute == null)
            {
                HighlightCell(worksheet.Cell(cellLetter + lineNumber));
            }
            else
            {
                worksheet.Cell(cellLetter + lineNumber).Value = customAttribute.ConstructorArguments.First().Value.ToString();
            }
        }

        private void FillCellCategory(IXLWorksheet worksheet, MethodInfo method, int lineNumber)
        {
            var cellLetter = GetLetterByNumber(columns.IndexOf("Category"));
            var categories = method.CustomAttributes.Where(a => a.AttributeType.Name.Contains("Category")).ToArray();

            if (categories == null)
            {
                HighlightCell(worksheet.Cell(cellLetter + lineNumber));

            }
            else
            {
                var categoriesValues = categories.Select(c => c.ConstructorArguments.First().Value).ToArray();
                worksheet.Cell(cellLetter + lineNumber).Value = string.Join(", ", categoriesValues);
            }

        }

        private void FillCellProperty(IXLWorksheet worksheet, MethodInfo method, int lineNumber)
        {
            var cellLetter = GetLetterByNumber(columns.IndexOf("Property"));
            var properties = method.CustomAttributes.Where(a => a.AttributeType.Name.Contains("Property")).ToList();

            List<string> propertyValues = new List<string>();

            if (properties.Count == 0)
            {
                HighlightCell(worksheet.Cell(cellLetter + lineNumber));
            }
            else
            {
                for (int i = 0; i < properties.Count; i++)
                {
                    var constructorArguments = properties[i].ConstructorArguments;
                    propertyValues.Add($"{constructorArguments[0].Value} -> {constructorArguments[1].Value}");
                }

                worksheet.Cell(cellLetter + lineNumber).Value = string.Join($"{Environment.NewLine}", propertyValues);
            }
        }

        private void AddSorting(IXLWorksheet worksheet)
        {
            worksheet.RangeUsed().CreateTable().Sort($"Test");
        }

        private void HighlightCell(IXLCell cell)
        {
            cell.Style.Fill.BackgroundColor = XLColor.Red;
        }

        private string GetLetterByNumber(int number)
        {
            return ((char)(number + 65)).ToString();
        }
    }
}
