Imports HybridShapeTypeLib
Imports INFITF
Imports MECMOD
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Inter

Namespace CatiaUnitTestProject
    Class App
        Public Shared CATIA As Application = Interaction.CreateObject(, "CATIA.APPLICATION")
    End Class
    <TestClass>
    Public Class UnitTest1

        <TestMethod>
        Sub TestSub()


            Dim partDocument1 As PartDocument

            partDocument1 = CATIA.ActiveDocument

Set part1 = partDocument1.Part

Set hybridShapeFactory1 = part1.HybridShapeFactory

Set hybridBodies1 = part1.HybridBodies

Set Selection = partDocument1.Selection

Set hybridBody1 = hybridBodies1.Item("Geometrical Set.1")

Set hybridShapes1 = hybridBody1.HybridShapes

Set hybridShapeLinePtPt1 = hybridShapes1.Item("Line.1")

Set reference1 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

Set hybridShapeSweepCircle1 = hybridShapeFactory1.AddNewSweepCircle(reference1)

HybridShapeSweepCircle.Mode = 6

            hybridShapeSweepCircle1.SmoothActivity = False

            hybridShapeSweepCircle1.GuideDeviationActivity = False

            hybridShapeSweepCircle1.SetRadius 1, 20.0

hybridShapeSweepCircle1.SetbackValue = 0.02

            hybridShapeSweepCircle1.FillTwistedAreas = 1

            hybridBody1.AppendHybridShape hybridShapeSweepCircle1

part1.InWorkObject = hybridShapeSweepCircle1

            part1.Update

            part1.Update

        End Sub
    End Class
End Namespace

