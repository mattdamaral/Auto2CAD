Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class CommonFunctions

    ' Draws a line in the model space
    Public Shared Sub DrawLine(startPoint As Point3d, endPoint As Point3d, colorIndex As Integer)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        ' Starts a transaction
        Using trans As Transaction = db.TransactionManager.StartTransaction()

            ' Opens the Block Table for reading
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

            ' Opens the Block Table Record for writing
            Using btr As BlockTableRecord = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                startPoint = startPoint
                endPoint = endPoint
                Dim line As Line = New Line(startPoint, endPoint)
                line.ColorIndex = colorIndex

                btr.AppendEntity(line)
                trans.AddNewlyCreatedDBObject(line, True)

            End Using

            trans.Commit()

        End Using
    End Sub

    ' -----------------------------------------------------------------------------------------------------------------------------------------------

    Public Shared Sub ChangeLayer(layerName As String)

        ' Accessess the database -------------------------------------------------------------------------------------------
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' Creates a transaction for the database ---------------------------------------------------------------------------
        Using trans As Transaction = doc.TransactionManager.StartTransaction()

            Dim lt As LayerTable
            Dim ltr As New LayerTableRecord
            Dim layerID As ObjectId

            ' Checks if layer exists
            Try
                ' If layer exists, get layerid
                lt = CType(trans.GetObject(db.LayerTableId, OpenMode.ForRead, True, True), LayerTable)
                layerID = lt.Item(layerName)

                ' If the layer was deleted, recover layer
                If layerID.IsErased Then
                    lt.UpgradeOpen()
                    lt.Item(layerName).GetObject(OpenMode.ForWrite, True, True).Erase(False)
                End If

            Catch ex As Exception
                ' If the layer doesn't exist, create it
                lt = db.LayerTableId.GetObject(OpenMode.ForWrite, True, True)
                ltr.Name = layerName
                lt.Add(ltr)
                ' Adds layer to db
                trans.AddNewlyCreatedDBObject(ltr, True)
                ' Recovers layerid of newly created layer
                lt = CType(trans.GetObject(db.LayerTableId, OpenMode.ForRead, False), LayerTable)
                layerID = lt.Item(layerName)
            End Try

            ' Sets layer as current
            db.Clayer = layerID

            trans.Commit()

        End Using

    End Sub

End Class
