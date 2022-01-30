using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPIPipeValue
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();

            string excelPath = @"G:\Рабочий стол\Повышение квалификации\BIM проектирование\Revit API\Задание 4.2 Вывод значений труб\pipes.xlsx";

            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Лист 1");

                int rowIndex = 0;
                foreach (var pipe in pipes)
                {
                    //используем предварительно созданный статический класс SheetEstc с методом SetCellValue
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.Name);
                    sheet.SetCellValue(rowIndex, columnIndex: 1, pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsValueString());
                    sheet.SetCellValue(rowIndex, columnIndex: 2, pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsValueString());
                    sheet.SetCellValue(rowIndex, columnIndex: 3, pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString());
                    rowIndex++;
                }

                workbook.Write(stream);
                workbook.Close();
            }

            //автоматическое открытие файла excel
            System.Diagnostics.Process.Start(excelPath);

            return Result.Succeeded;
        }
    }
}
