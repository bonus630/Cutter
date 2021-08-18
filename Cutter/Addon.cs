using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Corel.Interop.VGCore;

namespace Cutter
{
    public class Addon
    {
        private Application corelApp;
        private int index = 0;
        private readonly string alphaCounter = " ABCDEFGUIJKLMNOPQRSTUWVXYZ";
        public Addon(Application app)
        {
            this.corelApp = app;
        }

        public void test2()
        {
            ShapeRange sr = corelApp.ActiveSelectionRange;
            double oneMM = this.corelApp.ConvertUnits(0.1, cdrUnit.cdrMillimeter, corelApp.ActiveDocument.Unit);
            for (int i = 1; i <= sr.Count; i++)
            {
                Rect rect = sr[i].BoundingBox;
                this.corelApp.ActiveLayer.CreateEllipse2(rect.CenterX, rect.Top, oneMM);
            }


           // corelApp.MsgShow(oneMM);
        }
        public void test()
        {
            try
            {

                Shape guia = this.corelApp.ActiveDocument.ActivePage.FindShape("", cdrShapeType.cdrGuidelineShape);
                Shape rect = this.corelApp.ActiveShape;
                // corelApp.MsgShow(guia.Curve.Segments.Count.ToString());
                // bool ispoint = rect.BoundingBox.IsPointInside(rect.BoundingBox.CenterX, guia.PositionY);
                //System.Windows.MessageBox.Show(ispoint.ToString());
                Shape line = corelApp.ActiveVirtualLayer.CreateLineSegment(rect.LeftX - 1, guia.PositionY, rect.RightX + 1, guia.PositionY);
                Segments segs = rect.Curve.Segments;
                List<SegmentPoint> cps = new List<SegmentPoint>();
                for (int i = 1; i <= segs.Count; i++)
                {
                 //   corelApp.MsgShow(segs.Count.ToString());
                    Segment seg = segs[i];
                    CrossPoints cp = seg.GetIntersections(line.DisplayCurve.Segments[1]);
                    for (int j = 1; j <= cp.Count; j++)
                    {


                        double offset = 0;
                        if (seg.FindParamOffsetAtPoint(cp[j].PositionX, cp[j].PositionY, out offset))
                        {
                            try
                            {
                                seg.BreakApartAt(offset);
                            }
                            catch { }
                            //cps.Add(new SegmentPoint(seg.AbsoluteIndex, offset));
                            //   offset = seg.GetAbsoluteOffset(offset);
                         
                        }
                    }
                }
                line.Delete();
                //for (int i = 0; i < cps.Count; i++)
                //{
                //    Node node = rect.Curve.Segments. [cps[i].Segment+i].AddNodeAt(cps[i].Offset);
                //    node.BreakApart();
                //}
                rect.BreakApart();
            }
            catch (Exception e) { corelApp.MsgShow(e.Message); }

        }
    }
    public class SegmentPoint
    {
        public int Segment { get; protected set; }
        public double Offset { get; protected set; }
        public SegmentPoint(int segment,double offset)
        {
            Segment = segment;
            Offset = offset;
        }
    }
}
