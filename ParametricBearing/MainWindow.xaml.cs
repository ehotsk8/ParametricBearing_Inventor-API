using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Inventor;
using System.Diagnostics;

namespace ParametricBearing
{
    public partial class MainWindow : Window
    {
        CreateAssemblyClass assembly = new CreateAssemblyClass();
        private string[] selectNow = new string[10];
        private double d, D, B, r, Dw;
        public MainWindow()
        {
            InitializeComponent();
            DataGrido.ItemsSource = LoadCollectionData(selectNow);          
        }

        private void Create_Details_Click(object sender, RoutedEventArgs e)
        {
            UnitsTypeEnum[] LengthUnits = new UnitsTypeEnum[] { UnitsTypeEnum.kMillimeterLengthUnits,
                UnitsTypeEnum.kCentimeterLengthUnits,
                UnitsTypeEnum.kMeterLengthUnits };
            UnitsTypeEnum[] AngleUnits = new UnitsTypeEnum[] { UnitsTypeEnum.kRadianAngleUnits,
                UnitsTypeEnum.kGradAngleUnits };

            CreateDetails pDoc = new CreateDetails(d, D, B, r, Dw);
            PartDocument partDocument;
            partDocument = pDoc.CreatePartDocument("Внешнее кольцо подшипника", LengthUnits[0], AngleUnits[1]);
            partDocument.Views[1].GoHome();
            pDoc.Mirror_Obj(pDoc.Create_Revolve(pDoc.OutsideRing(), PartFeatureOperationEnum.kJoinOperation), partDocument);
            partDocument.SaveAs(@"C:\BearingDetails\Внешнее кольцо подшипника.ipt", false);
            partDocument = pDoc.CreatePartDocument("Внутреннее кольцо подшипника", LengthUnits[0], AngleUnits[1]);
            partDocument.Views[1].GoHome();
            pDoc.Mirror_Obj(pDoc.Create_Revolve(pDoc.InsideRing(), PartFeatureOperationEnum.kJoinOperation), partDocument);
            partDocument.SaveAs(@"C:\BearingDetails\Внутреннее кольцо подшипника.ipt", false);
            partDocument = pDoc.CreatePartDocument("Первый ряд шариков", LengthUnits[0], AngleUnits[1]);
            partDocument.Views[1].GoHome();
            pDoc.Create_Circular_Array(pDoc.Create_Revolve(pDoc.BallsRow1(), PartFeatureOperationEnum.kJoinOperation));
            partDocument.SaveAs(@"C:\BearingDetails\Первый ряд шариков.ipt", false);
            partDocument = pDoc.CreatePartDocument("Второй ряд шариков", LengthUnits[0], AngleUnits[1]);
            partDocument.Views[1].GoHome();
            pDoc.Create_Circular_Array(pDoc.Create_Revolve(pDoc.BallsRow2(), PartFeatureOperationEnum.kJoinOperation));
            partDocument.SaveAs(@"C:\BearingDetails\Второй ряд шариков.ipt", false);
            partDocument = pDoc.CreatePartDocument("Сепаратор1", LengthUnits[0], AngleUnits[1]);
            partDocument.Views[1].GoHome();
            pDoc.Separator(1);
            partDocument.SaveAs(@"C:\BearingDetails\Сепаратор1.ipt", false);
            partDocument = pDoc.CreatePartDocument("Сепаратор2", LengthUnits[0], AngleUnits[1]);
            partDocument.Views[1].GoHome();
            pDoc.Separator(-1);
            partDocument.SaveAs(@"C:\BearingDetails\Сепаратор2.ipt", false);
        }

        private void Create_Assembly_Click(object sender, RoutedEventArgs e)
        {
            assembly.Assembly();
        }

        private void Animator_Click(object sender, RoutedEventArgs e)
        {
            assembly.Anim_Play();
        }

        private List<Author> LoadCollectionData(string[] selectNow1)
        {
            List<Author> authors = new List<Author>();
            authors.Add(new Author() { d = 30, D = 62, B = 20, r = 1.5, Dw = 7.96, Выбраный_размер = selectNow1[1] });
            authors.Add(new Author() { d = 35, D = 72, B = 23, r = 2, Dw = 9.53, Выбраный_размер = selectNow1[2] });
            authors.Add(new Author() { d = 40, D = 80, B = 23, r = 2, Dw = 9.53, Выбраный_размер = selectNow1[3] });
            authors.Add(new Author() { d = 45, D = 85, B = 23, r = 2, Dw = 9.53, Выбраный_размер = selectNow1[4] });
            authors.Add(new Author() { d = 50, D = 90, B = 23, r = 2, Dw = 9.53, Выбраный_размер = selectNow1[5] });
            authors.Add(new Author() { d = 55, D = 100, B = 25, r = 2.5, Dw = 9.53, Выбраный_размер = selectNow1[6] });
            authors.Add(new Author() { d = 60, D = 110, B = 28, r = 2.5, Dw = 9.53, Выбраный_размер = selectNow1[7] });
            authors.Add(new Author() { d = 65, D = 120, B = 31, r = 2.5, Dw = 9.53, Выбраный_размер = selectNow1[8] });
            return authors;
        }

        private void ComboB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            for (int i = 0; i < 10; i++)
                selectNow[i] = "";
            selectNow[ComboB.SelectedIndex + 1] = "Выбран размер №" + (ComboB.SelectedIndex + 1);
            DataGrido.ItemsSource = LoadCollectionData(selectNow);
            d = LoadCollectionData(selectNow)[ComboB.SelectedIndex].d / 10;
            D = LoadCollectionData(selectNow)[ComboB.SelectedIndex].D / 10;
            B = LoadCollectionData(selectNow)[ComboB.SelectedIndex].B / 10;
            r = LoadCollectionData(selectNow)[ComboB.SelectedIndex].r / 10;
            Dw = LoadCollectionData(selectNow)[ComboB.SelectedIndex].Dw / 10;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Process.Start("inventor.exe");       
        }
    }

    public class Author
    {
        public double d { get; set; }
        public double D { get; set; }
        public double B { get; set; }
        public double r { get; set; }
        public double Dw { get; set; }
        public string Выбраный_размер { get; set; }
    }

    public class CreateDetails
    {
        double d, D, B, r, Dw, z;
        Inventor.Application InventorApplication;
        PartComponentDefinition partCompDef;
        TransientGeometry transGeom;
        SketchLine center_line;

        public CreateDetails(double dd, double DD, double BB, double rr, double DDw)
        {
            d = dd; D = DD; B = BB; r = rr; Dw = DDw;
        }

        public PartDocument CreatePartDocument(string DetailName, UnitsTypeEnum LengthUnits, UnitsTypeEnum AngleUnits)
        {
            InventorApplication = (Inventor.Application)
             System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
            PartDocument part = InventorApplication.Documents.Add(DocumentTypeEnum.kPartDocumentObject, "", true) as PartDocument;
            part.DisplayName = DetailName;
            part.UnitsOfMeasure.LengthUnits = LengthUnits;
            part.UnitsOfMeasure.AngleUnits = AngleUnits;
            partCompDef = part.ComponentDefinition;
            transGeom = InventorApplication.TransientGeometry;
            return part;
        }

        public PlanarSketch InsideRing()
        {
            PlanarSketch sketch = partCompDef.Sketches.Add(partCompDef.WorkPlanes[3]);
            SketchPoints points_ = sketch.SketchPoints;
            SketchArcs arcs_ = sketch.SketchArcs;
            SketchLines lines_ = sketch.SketchLines;
            points_.Add(transGeom.CreatePoint2d(0, d / 2));
            points_.Add(transGeom.CreatePoint2d(B / 2 - r, d / 2));
            points_.Add(transGeom.CreatePoint2d(B / 2 - r, d / 2 + r));
            points_.Add(transGeom.CreatePoint2d(B / 2, d / 2 + r));
            points_.Add(transGeom.CreatePoint2d(B / 2, d / 2 + r + 0.25 * d / 3));
            points_.Add(transGeom.CreatePoint2d(B / 2 - 0.15 * B / 2, d / 2 + r + 0.40 * d / 3));
            points_.Add(transGeom.CreatePoint2d(B / 2 - 0.4 * B / 2, d / 2 + r + 0.40 * d / 3));
            points_.Add(transGeom.CreatePoint2d(0.2 * B / 2, d / 2 + r + 0.40 * d / 3));
            points_.Add(transGeom.CreatePoint2d(0, d / 2 + r + 0.40 * d / 3));
            points_.Add(transGeom.CreatePoint2d(Dw / 2, (D / 2 - d / 2) / 2 + d / 2));
            lines_.AddByTwoPoints(points_[1], points_[2]);
            arcs_.AddByCenterStartEndPoint(points_[3], points_[2], points_[4], true);
            for (int i = 4; i < 7; i++)
                lines_.AddByTwoPoints(points_[i], points_[i + 1]);
            lines_.AddByTwoPoints(points_[8], points_[9]);
            lines_.AddByTwoPoints(points_[9], points_[1]);
            arcs_.AddByCenterStartEndPoint(points_[10], points_[7], points_[8], false);
            sketch.GeometricConstraints.AddGround(sketch.SketchEntities[10]);
            arcs_[2].StartSketchPoint.Merge(points_[8]);
            arcs_[2].CenterSketchPoint.Merge(points_[10]);
            for (int i = 1; i < 6; i++)
                sketch.GeometricConstraints.AddGround(sketch.SketchEntities[i]);
            SketchEntity entity = sketch.SketchEntities[19];
            sketch.DimensionConstraints.AddRadius(entity, points_[1].Geometry);
            sketch.DimensionConstraints[1].Parameter.Value = Dw / 2;
            center_line = sketch.SketchLines.AddByTwoPoints(transGeom.CreatePoint2d(-1, 0), transGeom.CreatePoint2d(1, 0));
            return sketch;
        }

        //Скетч внешнего кольца
        public PlanarSketch OutsideRing()
        {
            PlanarSketch sketch = partCompDef.Sketches.Add(partCompDef.WorkPlanes[3]);
            SketchPoints points_ = sketch.SketchPoints;
            SketchArcs arcs_ = sketch.SketchArcs;
            SketchLines lines_ = sketch.SketchLines;
            points_.Add(transGeom.CreatePoint2d(0, D / 2));
            points_.Add(transGeom.CreatePoint2d(B / 2 - r, D / 2));
            points_.Add(transGeom.CreatePoint2d(B / 2 - r, D / 2 - r));
            points_.Add(transGeom.CreatePoint2d(B / 2, D / 2 - r));
            points_.Add(transGeom.CreatePoint2d(B / 2, D / 2 - r - 0.3 * D / 6.2));
            points_.Add(transGeom.CreatePoint2d(B / 2 - 0.1 * B / 2, D / 2 - r - 0.4 * D / 6.2));
            points_.Add(transGeom.CreatePoint2d(Dw / 2, (D / 2 - d / 2) / 2 + d / 2 + Dw / 2));
            points_.Add(transGeom.CreatePoint2d(-Dw / 2 , (D / 2 - d / 2) / 2 + d / 2 + Dw / 2));
            lines_.AddByTwoPoints(points_[1], points_[2]);
            arcs_.AddByCenterStartEndPoint(points_[3], points_[2], points_[4], false);
            for (int i = 4; i < 6; i++)
                lines_.AddByTwoPoints(points_[i], points_[i + 1]);
            arcs_.AddByCenterStartEndPoint(transGeom.CreatePoint2d(0, 0), points_[8], points_[6], false);
            lines_[3].EndSketchPoint.Merge(arcs_[2].StartSketchPoint);
            lines_.AddByTwoPoints(arcs_[2].EndSketchPoint, points_[1]);
            center_line = sketch.SketchLines.AddByTwoPoints(transGeom.CreatePoint2d(-1, 0), transGeom.CreatePoint2d(1, 0));
            return sketch;
        }

        //Скетч шариков
        public PlanarSketch BallsRow1()
        {
            PlanarSketch sketch = partCompDef.Sketches.Add(partCompDef.WorkPlanes[3]);
            SketchPoints points_ = sketch.SketchPoints;
            points_.Add(transGeom.CreatePoint2d(0, (D / 2 - d / 2) / 2 + d / 2 + Dw / 2));
            points_.Add(transGeom.CreatePoint2d(0, (D / 2 - d / 2) / 2 + d / 2));
            points_.Add(transGeom.CreatePoint2d(0, (D / 2 - d / 2) / 2 + d / 2 - Dw / 2));
            SketchArcs arcs_ = sketch.SketchArcs;
            arcs_.AddByCenterStartEndPoint(points_[2], points_[1], points_[3], true);
            Point2d center_point = transGeom.CreatePoint2d(Dw / 2 , (D / 2 - d / 2) / 2 + d / 2);
            SketchPoint points = sketch.SketchPoints.Add(transGeom.CreatePoint2d(arcs_[1].Radius + Dw / 2 , (D / 2 - d / 2) / 2 + d / 2));
            SketchPoint points2 = sketch.SketchPoints.Add(transGeom.CreatePoint2d(-arcs_[1].Radius + Dw / 2 , (D / 2 - d / 2) / 2 + d / 2));
            center_line = sketch.SketchLines.AddByTwoPoints(points, points2);
            SketchArc arc = sketch.SketchArcs.AddByCenterStartEndPoint(center_point, points, points2, true);
            return sketch;
        }

        //Скетч шариков
        public PlanarSketch BallsRow2()
        {
            PlanarSketch sketch = partCompDef.Sketches.Add(partCompDef.WorkPlanes[3]);
            SketchPoints points_ = sketch.SketchPoints;
            points_.Add(transGeom.CreatePoint2d(0, (D / 2 - d / 2) / 2 + d / 2 + Dw / 2 ));
            points_.Add(transGeom.CreatePoint2d(0, (D / 2 - d / 2) / 2 + d / 2));
            points_.Add(transGeom.CreatePoint2d(0, (D / 2 - d / 2) / 2 + d / 2 - Dw / 2 ));
            SketchArcs arcs_ = sketch.SketchArcs;
            arcs_.AddByCenterStartEndPoint(points_[2], points_[1], points_[3], true);
            Point2d center_point = transGeom.CreatePoint2d(-Dw / 2 , (D / 2 - d / 2) / 2 + d / 2);
            SketchPoint points = sketch.SketchPoints.Add(transGeom.CreatePoint2d(arcs_[1].Radius - Dw / 2 , (D / 2 - d / 2) / 2 + d / 2));
            SketchPoint points2 = sketch.SketchPoints.Add(transGeom.CreatePoint2d(-arcs_[1].Radius - Dw / 2 , (D / 2 - d / 2) / 2 + d / 2));
            center_line = sketch.SketchLines.AddByTwoPoints(points, points2);
            SketchArc arc = sketch.SketchArcs.AddByCenterStartEndPoint(center_point, points, points2, true);
            return sketch;
        }

        public PlanarSketch Separator(double direction)
        {
            PlanarSketch sketch = partCompDef.Sketches.Add(partCompDef.WorkPlanes[3]);
            SketchPoints points_ = sketch.SketchPoints;
            SketchArcs arcs_ = sketch.SketchArcs;
            SketchLines lines_ = sketch.SketchLines;
            Point2d center_point = transGeom.CreatePoint2d(Dw / 2 * direction, (D / 2 - d / 2) / 2 + d / 2);
            SketchPoint points = sketch.SketchPoints.Add(transGeom.CreatePoint2d(Dw  * direction, (D / 2 - d / 2) / 2 + d / 2));
            SketchPoint points2 = sketch.SketchPoints.Add(transGeom.CreatePoint2d(0, (D / 2 - d / 2) / 2 + d / 2));
            center_line = sketch.SketchLines.AddByTwoPoints(points, points2);
            SketchArc arc = sketch.SketchArcs.AddByCenterStartEndPoint(center_point, points, points2, true);
            RevolveFeature revolve1 = Create_Revolve(sketch, PartFeatureOperationEnum.kNewBodyOperation);
            Create_Circular_Array(revolve1);
            sketch = partCompDef.Sketches.Add(partCompDef.WorkPlanes[3]);
            points_ = sketch.SketchPoints;
            lines_ = sketch.SketchLines;
            points_.Add(transGeom.CreatePoint2d((Dw / 2 + 0.25)  * direction, (D / 2 - d / 2) / 2 + d / 2 + 0.15));
            points_.Add(transGeom.CreatePoint2d(0, (D / 2 - d / 2) / 2 + d / 2 + 0.3));
            points_.Add(transGeom.CreatePoint2d(0, (D / 2 - d / 2) / 2 + d / 2 - 0.15));
            points_.Add(transGeom.CreatePoint2d(Dw / 2  * direction, (D / 2 - d / 2) / 2 + d / 2 - 0.1));
            for (int i = 1; i < 4; i++)
                lines_.AddByTwoPoints(points_[i], points_[i + 1]);
            ObjectCollection obj_collection = InventorApplication.TransientObjects.CreateObjectCollection();
            obj_collection.Add(sketch.SketchEntities[5]);
            obj_collection.Add(sketch.SketchEntities[6]);
            obj_collection.Add(sketch.SketchEntities[7]);
            sketch.OffsetSketchEntitiesUsingDistance(obj_collection, 0.1  * direction, false);
            lines_.AddByTwoPoints(sketch.SketchLines[1].StartSketchPoint, sketch.SketchLines[4].StartSketchPoint);
            lines_.AddByTwoPoints(sketch.SketchLines[3].EndSketchPoint, sketch.SketchLines[6].EndSketchPoint);
            center_line = sketch.SketchLines.AddByTwoPoints(transGeom.CreatePoint2d(-1, 0), transGeom.CreatePoint2d(1, 0));
            RevolveFeature revolve2 = Create_Revolve(sketch, PartFeatureOperationEnum.kNewBodyOperation);
            obj_collection = InventorApplication.TransientObjects.CreateObjectCollection();
            obj_collection.Add(revolve1.SurfaceBodies[1]);
            CombineFeature combine = partCompDef.Features.CombineFeatures.Add(revolve2.SurfaceBodies[1], obj_collection, PartFeatureOperationEnum.kCutOperation);
            return sketch;
        }

        //Массив по кругу
        public void Create_Circular_Array(RevolveFeature revolve)
        {
            ObjectCollection obj_collection = InventorApplication.TransientObjects.CreateObjectCollection();
            obj_collection.Add(revolve);
            WorkAxis XAxis = partCompDef.WorkAxes[1]; 
            CircularPatternFeature pattern = partCompDef.Features.CircularPatternFeatures.Add(obj_collection, XAxis, false, 10, "360 grad", false, PatternComputeTypeEnum.kIdenticalCompute);
        }

        //Создание эл-та кручение
        public RevolveFeature Create_Revolve(PlanarSketch sketch, PartFeatureOperationEnum enumer)
        {
            Profile profile = sketch.Profiles.AddForSolid();
            RevolveFeature revolve = partCompDef.Features.RevolveFeatures.AddFull(profile, center_line, enumer);
            return revolve;
        }

        //Зеркальное отражение детали
        public void Mirror_Obj(RevolveFeature feature, PartDocument part)
        {
            PartComponentDefinition partCompDef = part.ComponentDefinition;
            PlanarSketch sketch = partCompDef.Sketches.Add(partCompDef.WorkPlanes[1]);
            sketch.Visible = false;
            TransientGeometry transGeom = InventorApplication.TransientGeometry;
            SketchLine line_1 = sketch.SketchLines.AddByTwoPoints(transGeom.CreatePoint2d(-6, -5), transGeom.CreatePoint2d(-6, 5));
            SketchLine line_2 = sketch.SketchLines.AddByTwoPoints(transGeom.CreatePoint2d(6, -5), transGeom.CreatePoint2d(6, 5));
            WorkPlane wp = partCompDef.WorkPlanes.AddByTwoLines(line_1, line_2, true);
            ObjectCollection obj_collection = InventorApplication.TransientObjects.CreateObjectCollection();
            obj_collection.Add(feature);
            MirrorFeature mirror = partCompDef.Features.MirrorFeatures.Add(obj_collection, wp, false, PatternComputeTypeEnum.kAdjustToModelCompute);
        }
    }
}

//Класс создание сборки
public class CreateAssemblyClass
{
    private AssemblyComponentDefinition assDefinition;
    private AssemblyConstraints constaints;
    private AssemblyJoints joints;

    public void Assembly()
    {
        Inventor.Application InventorApplication = (Inventor.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
        AssemblyDocument assemblyDoc = InventorApplication.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, "", true) as AssemblyDocument;
        assDefinition = assemblyDoc.ComponentDefinition;
        Inventor.Matrix matrix = InventorApplication.TransientGeometry.CreateMatrix();
        matrix.SetTranslation(InventorApplication.TransientGeometry.CreateVector(0, 0, 0));
        string[] pathName = new string[10];
        pathName[0] = @"C:\BearingDetails\Внешнее кольцо подшипника.ipt";
        pathName[1] = @"C:\BearingDetails\Внутреннее кольцо подшипника.ipt";
        pathName[2] = @"C:\BearingDetails\Первый ряд шариков.ipt";
        pathName[3] = @"C:\BearingDetails\Второй ряд шариков.ipt";
        pathName[4] = @"C:\BearingDetails\Сепаратор1.ipt";
        pathName[5] = @"C:\BearingDetails\Сепаратор2.ipt";
        for (int i = 0; i < 6; i++)
        { assDefinition.Occurrences.Add(pathName[i], matrix); }
        Rotate();
        for (int i = 1; i < 4; i++)
            Tangent(i, 2, 3, 7);
        Tangent(1, 5, 3, 1);
        Tangent(3, 5, 3, 5);
        Tangent(5, 5, 3, 9);
        //Tangent(1, 6, 4, 1);
        //Tangent(3, 6, 4, 5);
        //Tangent(5, 6, 4, 9);
        //for (int i = 2; i < 5; i++)
        //    Tangent(i, 2, 4, 1);
        assemblyDoc.SaveAs(@"C:\BearingDetails\Сборка_подшипника.iam", false);
    }

    //Зависимость касательной 
    public void Tangent(int i, int occurence_1, int occurence_2, int faces)
    {
        Face face1 = assDefinition.Occurrences[occurence_2].SurfaceBodies[1].Faces[i * 2];
        Face face2 = assDefinition.Occurrences[occurence_1].SurfaceBodies[1].Faces[faces];
        constaints = assDefinition.Constraints;
        constaints.AddTangentConstraint(face1, face2, true, 0);
    }

    //Соединение кручение
    public void Rotate()
    {
        GeometryIntent OriginOne = assDefinition.CreateGeometryIntent(assDefinition.Occurrences[1].SurfaceBodies[1].Faces[2].Edges[1]);
        GeometryIntent OriginTwo = assDefinition.CreateGeometryIntent(assDefinition.Occurrences[2].SurfaceBodies[1].Faces[3].Edges[2]);
        AssemblyJointDefinition jointDef = assDefinition.Joints.CreateAssemblyJointDefinition(AssemblyJointTypeEnum.kRotationalJointType, OriginTwo, OriginOne);
        joints = assDefinition.Joints;
        joints.Add(jointDef);
    }

    //Зависимость вставка
    public void Insert()
    {
        Face face1 = assDefinition.Occurrences[2].SurfaceBodies[1].Faces[3];
        Face face2 = assDefinition.Occurrences[1].SurfaceBodies[1].Faces[2];
        constaints = assDefinition.Constraints;
        constaints.AddInsertConstraint(face1, face2, false, 0);
    }

    //Анимация соединия кручение
    public void Anim_Play()
    {
        joints[1].DriveSettings.SetIncrement(IncrementTypeEnum.kAmountOfValueIncrement, "20 deg");
        joints[1].DriveSettings.StartValue = "0 deg";
        joints[1].DriveSettings.EndValue = "360 deg";
        joints[1].DriveSettings.PlayForward();
    }
}





