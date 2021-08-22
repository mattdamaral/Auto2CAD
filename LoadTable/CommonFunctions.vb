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

End Class
