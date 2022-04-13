using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePlugin
{
    [Transaction(TransactionMode.Manual)]

    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document; //обращаемся к основному открытому файлу
            Document ovDoc = arDoc.Application.Documents
                .OfType<Document>()
                .Where(x=>x.Title.Contains("ОВ"))
                .FirstOrDefault(); //обращаемся к файлу, содержащему в названии "ОВ", связанному с основным
            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка","Не обнаружен файл ОВ");
                return Result.Cancelled;
            }

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc) //для проверки загружено ли семейство с отверстием в файл АР
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстие"))
                .FirstOrDefault();
            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка","Не обнаружено семейство \"Отверстие\"");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc) //находм все воздуховоды в список
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(ovDoc) //находм все трубы в список
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();

            View3D view3D = new FilteredElementCollector(arDoc) //ищем есть ли у нас 3Д вид, т.к. метод поиска пересечения работает только на них
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate) //отсекаем шаблоны 3Д видов
                .FirstOrDefault();
            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не обнаружен 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);
            // создали экземпляр объекта - пересечения

            Transaction transaction0 = new Transaction(arDoc); //вспомогательная транзакция для аквтивации семейства (если оно не активно в модели)
            transaction0.Start("Активация семейства отверстия");

            if (!familySymbol.IsActive)
            {
                familySymbol.Activate();
            }
            transaction0.Commit();

            Transaction transaction1 = new Transaction(arDoc);
            transaction1.Start("Расстановка отверстий для воздуховодов");          

            foreach (Duct d in ducts) //для каждого воздуховода из коллекции
            {
                Line curve = (d.Location as LocationCurve).Curve as Line; //берем образующую кривую как линию
                XYZ point = curve.GetEndPoint(0); //берем исходную (первую) точку линии
                XYZ direction = curve.Direction; //берем направление линии

                List<ReferenceWithContext> intersections =  referenceIntersector.Find(point, direction) //получаем набор пересечений
                    .Where(x=>x.Proximity<=curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer()) //это дополнительно взят класс (ниже) для фильтрации точек, принадлежащих одной стене
                    .ToList();

                foreach (ReferenceWithContext r in intersections)
                {
                    double proximity = r.Proximity; //берем расстояние до пересечения
                    Reference reference = r.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall; //берем стену которую пересекает...сложно
                    Level level = arDoc.GetElement(wall.LevelId) as Level; //берем уровень стены, которую пересекаем
                    XYZ pointHole = point + (direction * proximity); //вот так просто находится точка персечения воздуховода и стены

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    //вставили отверстие в стену
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(d.Diameter); //устанавливаем ширину отверстия по диамету воздуховода
                    height.Set(d.Diameter); //устанавливаем высоту отверстия по диамету воздуховода
                }
            }
            transaction1.Commit();

            Transaction transaction2 = new Transaction(arDoc);
            transaction2.Start("Расстановка отверстий для труб");

            foreach (Pipe p in pipes) //для каждого воздуховода из коллекции
            {
                Line curve = (p.Location as LocationCurve).Curve as Line; //берем образующую кривую как линию
                XYZ point = curve.GetEndPoint(0); //берем исходную (первую) точку линии
                XYZ direction = curve.Direction; //берем направление линии

                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction) //получаем набор пересечений
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer()) //это дополнительно взят класс (ниже) для фильтрации точек, принадлежащих одной стене
                    .ToList();

                foreach (ReferenceWithContext r in intersections)
                {
                    double proximity = r.Proximity; //берем расстояние до пересечения
                    Reference reference = r.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall; //берем стену которую пересекает...сложно
                    Level level = arDoc.GetElement(wall.LevelId) as Level; //берем уровень стены, которую пересекаем
                    XYZ pointHole = point + (direction * proximity); //вот так просто находится точка персечения трубы и стены

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    //вставили отверстие в стену
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(p.Diameter); //устанавливаем ширину отверстия по диамету воздуховода
                    height.Set(p.Diameter); //устанавливаем высоту отверстия по диамету воздуховода
                }
            }
            transaction2.Commit();

            return Result.Succeeded;
        }
    }

    public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        //ВЗЯЛИ КЛАСС ЦЕЛИКОМ С ФОРУМА
    {
        public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (ReferenceEquals(null, x)) return false;
            if (ReferenceEquals(null, y)) return false;

            var xReference = x.GetReference();

            var yReference = y.GetReference();

            return xReference.LinkedElementId == yReference.LinkedElementId
                       && xReference.ElementId == yReference.ElementId; //определяет ели у обоих элементов одинаковый ElementID (точки на одной стене), то возвращает true
        }

        public int GetHashCode(ReferenceWithContext obj)
        {
            var reference = obj.GetReference();

            unchecked
            {
                return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
            }
        }
    }
}
