Attribute VB_Name = "Inventor_VBA_01"
Option Explicit

Public Sub Process()

    Dim n As Integer

    'Set part Document
    Dim partDoc As Inventor.PartDocument
    Set partDoc = ThisDocument
    
    'Manage a feaure
    Dim extrude As ExtrudeFeature
    Set extrude = partDoc.ComponentDefinition.Features.ExtrudeFeatures.Item("ExtrusionName1")
    extrude.Parameters(3).Value = 2
    For n = 1 To extrude.Parameters.Count
        'Debug.Print (extrude.Parameters(n).Name + " -> " + CStr(extrude.Parameters(n).Value))
    Next

    Dim fillet As FilletFeature
    Set fillet = partDoc.ComponentDefinition.Features.FilletFeatures.Item("Fillet1")
    fillet.Parameters(1).Value = 1

    'Manage a Sketch
    Dim sketch2 As Sketch
    Set sketch2 = partDoc.ComponentDefinition.Sketches.Item("sketch2")
    sketch2.SketchCircles(1).Radius = 2
    
    Dim PointOrg As Point2d
    Set PointOrg = ThisApplication.TransientGeometry.CreatePoint2d
    PointOrg.X = 0
    PointOrg.Y = 0
    
    Dim centre As SketchPoint
    Set centre = sketch2.SketchPoints.Add(PointOrg)
    
    Dim circle1 As SketchCircle
    'Set circle1 = sketch2.SketchCircles.AddByCenterRadius(centre, ThisDocument.ComponentDefinition.Parameters.UserParameters.Item("u2").Value)
    Set circle1 = sketch2.SketchCircles.AddByCenterRadius(centre, 6.5)
    
    'Manage Parameters
    Call ThisDocument.ComponentDefinition.Parameters.UserParameters.AddByValue("Dist1", 67, kMillimeterLengthUnits)
    Debug.Print ThisDocument.ComponentDefinition.Parameters.ModelParameters.Item("d1").Value
    Debug.Print ThisDocument.ComponentDefinition.Parameters.UserParameters.Item("u2").Value

    'Udate
    partDoc.Update
   
    
End Sub

