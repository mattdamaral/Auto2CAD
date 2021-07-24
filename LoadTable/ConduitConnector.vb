Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class ConduitConnector

    <CommandMethod("CC")>
    Public Sub ConnectConduit()

        ' Accessess the database -------------------------------------------------------------------------------------------
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' Creates a transaction for the database ---------------------------------------------------------------------------
        Using trans As Transaction = doc.TransactionManager.StartTransaction()

            ' Filter for the selection
            Dim tv As TypedValue() = New TypedValue(0) {}
            tv.SetValue(New TypedValue(CInt(DxfCode.Start), "INSERT"), 0)
            Dim filter As SelectionFilter = New SelectionFilter(tv)

            ' Selects objects using filter
            Dim psr As PromptSelectionResult = ed.GetSelection(filter)

            ' Checks if selection is ok
            If psr.Status = PromptStatus.OK Then

                Dim ss As SelectionSet = psr.Value

                Dim pos1X As Double = 0
                Dim pos1Y As Double = 0
                Dim foundPos1 As Boolean = False
                Dim pos2X As Double = 0
                Dim pos2Y As Double = 0
                Dim foundPos2 As Boolean = False

                ' Repeats code for every item selected
                For Each so As SelectedObject In ss

                    ' Blocks parameters
                    Dim br As BlockReference = CType(trans.GetObject(so.ObjectId, OpenMode.ForRead), BlockReference)
                    'Dim ac As AttributeCollection = br.AttributeCollection
                    Dim properties As DynamicBlockReferencePropertyCollection = br.DynamicBlockReferencePropertyCollection

                    ' Fills the attribute list (attList) with attribute references
                    For Each prop As DynamicBlockReferenceProperty In properties

                        'MsgBox(prop.PropertyName.ToString)
                        Dim propName = prop.PropertyName

                        'Dim attR As AttributeReference = CType(trans.GetObject(oid, OpenMode.ForRead), AttributeReference)  ' Stores the attribute
                        If foundPos1 = False Then
                            Select Case propName
                                Case propName.Contains("CON_ELET_0 X")
                                    Double.TryParse(propName, pos1X)
                                Case propName.Contains("CON_ELET_0 Y")
                                    Double.TryParse(propName, pos1Y)
                                    foundPos1 = True
                            End Select
                        Else
                            If foundPos2 = False Then
                                Select Case propName
                                    Case propName.Contains("CON_ELET_0 X")
                                        Double.TryParse(propName, pos2X)
                                    Case propName.Contains("CON_ELET_0 Y")
                                        Double.TryParse(propName, pos2Y)
                                        foundPos2 = True
                                End Select
                            End If
                        End If

                    Next

                Next

                If foundPos1 = True And foundPos2 = True Then

                    Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                    Dim btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    Dim point1 As Point2d = New Point2d(pos1X, pos1Y)
                    Dim point2 As Point2d = New Point2d(pos2X, pos2Y)
                    Dim arc As CircularArc2d = New CircularArc2d(point1, New Point2d(pos1X - pos2X + 5.77, pos1Y - pos2Y + 5.77), point2)

                    'Dim 

                    'btr.AppendEntity(arc)
                    'trans.AddNewlyCreatedDBObject(arc, True)

                    trans.Commit()

                End If

            End If

        End Using

    End Sub

End Class
