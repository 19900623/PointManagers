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
using INFITF;
using MECMOD;
using HybridShapeTypeLib;
using NPOI;
using NPOI.Util;
using NPOI.XSSF;
using NPOI.HSSF;
using NPOI.SS;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;
using System.Reflection;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using AnnotationTypeLib;
using System.ComponentModel;

namespace PointManager_CATIA
{

    public partial class MainWindow : System.Windows.Window
    {
        private HSSFWorkbook OpenFileXLS(string FileName)
        {
            try
            {
                var inputFile = new FileStream(FileName, FileMode.Open, FileAccess.Read);
                var Doc = new HSSFWorkbook(inputFile);
                return Doc;
            }
            catch (IOException)
            {
                return null;
            }
        }

        private XSSFWorkbook OpenFileXLSX(string FileName)
        {
            try
            {
                var inputFile = new FileStream(FileName, FileMode.Open, FileAccess.Read);
                var Doc = new XSSFWorkbook(inputFile);
                return Doc;
            }
            catch (IOException)
            {
                return null;
            }

        }

        BackgroundWorker ImportWorker;

        private void DrawPoints(IWorkbook Doc)
        {
            AllPoints.Clear();
            int errcount = 0;
            ISheet Sheet = Doc.GetSheetAt(0);
            for (int i = 1; i < Sheet.PhysicalNumberOfRows; i++)
            {
                if (Sheet.GetRow(i) != null)
                {
                    try
                    {
                        IRow Row = Sheet.GetRow(i);
                        string name = Row.GetCell(0).ToString();
                        string badsymbol = @"'";
                        double X = Convert.ToDouble(Row.GetCell(1).ToString().Replace('.', ',').Replace(badsymbol, "").Replace(" ",""));
                        double Y = Convert.ToDouble(Row.GetCell(2).ToString().Replace('.', ',').Replace(badsymbol, "").Replace(" ", ""));
                        double Z = Convert.ToDouble(Row.GetCell(3).ToString().Replace('.', ',').Replace(badsymbol, "").Replace(" ", ""));
                        NodePoint iNode = new NodePoint(name, X, Y, Z, null);
                        AllPoints.Add(iNode);
                    }
                    catch
                    {
                        errcount++;
                    }
                }
            }
            ////
            CATIA = (INFITF.Application)Marshal.GetActiveObject("Catia.Application");
            var partDoc = CATIA.ActiveDocument as PartDocument;
            if (partDoc != null)
            {
                var Part = partDoc.Part;
                var HBodies = Part.HybridBodies;
                var HBody = HBodies.Add();
                HBody.set_Name("Import Points_" + AllPoints.Count.ToString());
                Part.Update();
                var ShapeFactory = Part.HybridShapeFactory as HybridShapeFactory;
                foreach (NodePoint iNode in AllPoints)
                {
                    var Point = ShapeFactory.AddNewPointCoord(iNode.X, iNode.Y, iNode.Z);
                    Point.set_Name(iNode.Name);
                    iNode.link = Point as HybridShapeTypeLib.Point;
                    HBody.AppendHybridShape(Point);
                }
                Part.Update();
            }
            else
            {
                throw new InvalidOperationException("Document not found");
            }

        }
         

        private void SetAnnotations(string Type)
        {
            CATIA = (INFITF.Application)Marshal.GetActiveObject("Catia.Application");
            var partDoc = CATIA.ActiveDocument as PartDocument;
            var Part = partDoc.Part;
            var AnnSets = Part.AnnotationSets as AnnotationSets;
            var AnnSet = AnnSets.Add("ISO_3D") as AnnotationSet;
            var AnnFactory = AnnSet.AnnotationFactory as AnnotationFactory;
            var HBodies = Part.HybridBodies;
            var HBody = HBodies.Item("Import Points_" + AllPoints.Count.ToString());
            var HShapes = HBody.HybridShapes as HybridShapes;
            var Surfaces = Part.UserSurfaces as UserSurfaces;
            double Z0 = AllPoints[0].Z;

            var myWindow = CATIA.ActiveWindow as SpecsAndGeomWindow;
            var myViewer3D = myWindow.ActiveViewer as Viewer3D;
            var myViewPoint = myViewer3D.Viewpoint3D as Viewpoint3D;
            var myCams = partDoc.Cameras as Cameras;

            foreach (NodePoint iNode in AllPoints)
            {
                var Point = HShapes.Item(iNode.Name) as HybridShapePointCoord;
                var Ref = Part.CreateReferenceFromObject(Point) as Reference;
                var Surf = Surfaces.Generate(Ref);
                if (Type == "Windshield")
                {
                    var myCam = myCams.Item(2) as Camera3D;
                    myViewPoint = myCam.Viewpoint3D;
                    myViewer3D.Viewpoint3D = myViewPoint;
                }
                if (Type == "Backlite")
                {
                    var myCam = myCams.Item(3) as Camera3D;
                    myViewPoint = myCam.Viewpoint3D;
                    myViewer3D.Viewpoint3D = myViewPoint;
                }
                if (Type == "Left Door")
                {
                    var myCam = myCams.Item(5) as Camera3D;
                    myViewPoint = myCam.Viewpoint3D;
                    myViewer3D.Viewpoint3D = myViewPoint;
                }
                if (Type == "Right Door")
                {
                    var myCam = myCams.Item(4) as Camera3D;
                    myViewPoint = myCam.Viewpoint3D;
                    myViewer3D.Viewpoint3D = myViewPoint;
                }
                var Ann = AnnFactory.CreateEvoluateText(Surf, iNode.X + 15, iNode.Y + 15, Z0 - iNode.Z, true);
                Ann.Text().set_Text(iNode.Name);
            }
            Part.Update();
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            ImportWorker = new BackgroundWorker();
            ImportWorker.DoWork += ImportWorker_DoWork;
            ImportWorker.RunWorkerCompleted += ImportWorker_RunWorkerCompleted;

            OpenFileDialog myDialog = new OpenFileDialog();
            string FileName = "";
            myDialog.Filter = "Документы Excel (*.xlsx)|**.XLSX";
            if (myDialog.ShowDialog() == true)
            {
                FileName = myDialog.FileName;
                ExportButton.Content = "Подождите...";
                ImportButton.Content = "Подождите...";
                ImportButton.IsEnabled = false;
                ExportButton.IsEnabled = false;

                var p = new Iparams();
                p.name = FileName;
                p.annotate = checkBox.IsChecked;
                p.type = ViewTypeCombo.SelectedValue.ToString();
                ImportWorker.RunWorkerAsync(p);
            }
        }

        public class Iparams
        {
            public string name;
            public string type;
            public bool? annotate;
        }


        private void ImportWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ExportButton.Content = "Выполнить";
            ImportButton.Content = "Открыть файл";
            ImportButton.IsEnabled = true;
            ExportButton.IsEnabled = true;
            var res = e.Result as string;
            if (res == "ok")
            {
                MessageBox.Show("Все точки проставлены.", "Готово!", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("Документ не найден. Проставление точек невозможно при работе с CATProduct. Откройте нужную деталь (CATPart) отдельно.", "Упс!", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            
        }

        private void ImportWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var p = e.Argument as Iparams;
            string FileName = p.name;

            var inputFile = new FileStream(FileName, FileMode.Open, FileAccess.Read);
            if (FileName.ToLower().EndsWith(".xlsx"))
            {
                XSSFWorkbook inbook = OpenFileXLSX(FileName);
                try
                {
                    DrawPoints(inbook);
                    if (p.annotate == true)
                    {
                        SetAnnotations(p.type);
                    }
                    e.Result = "ok";
                }
                catch(InvalidOperationException ex)
                {
                    e.Result = ex.Message;
                }
            }
            if (FileName.ToLower().EndsWith(".xls"))
            {
                HSSFWorkbook inbook = OpenFileXLS(FileName);
                try
                {
                    DrawPoints(inbook);
                    if (p.annotate == false)
                    {
                        SetAnnotations(p.type);
                    }
                    e.Result = "ok";
                }
                catch (InvalidOperationException ex)
                {
                    e.Result = ex.Message;
                }
            }
        }
    }
}

