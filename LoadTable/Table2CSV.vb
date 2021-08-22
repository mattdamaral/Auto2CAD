'Imports Autodesk.AutoCAD.ApplicationServices
'Imports Autodesk.AutoCAD.DatabaseServices
'Imports Autodesk.AutoCAD.Runtime
'Imports Autodesk.AutoCAD.Geometry
'Imports Autodesk.AutoCAD.EditorInput

'Imports System
'Imports System.IO

''Imports System.Text
''Imports Autodesk.AutoCAD.ApplicationServices
''Imports Autodesk.AutoCAD.DatabaseServices
''Imports Autodesk.AutoCAD.EditorInput
''Imports Autodesk.AutoCAD.Runtime
''Imports System
''Imports System.IO



''Namespace TableExportUnicode

''{

'Public Class Table2CSV

'    '[CommandMethod("T2CSV", CommandFlags.UsePickSet)]
'    <CommandMethod("T2CSV")>
'    Public Sub ExportTableToUnicode()

'        Dim doc = Application.DocumentManager.MdiActiveDocument
'        Dim db = doc.Database
'        Dim ed = doc.Editor

'        Dim tbId = ObjectId.Null 

'        'Check the pickfirst selection for a single Table object

'        Dim psr = ed.GetSelection()

'        If (psr.Status <> PromptStatus.OK) Then
'            Return
'        End If

'        If (psr.Value.Count = 1) Then

'            Dim selId = psr.Value(0).ObjectId

'            'If selId.ObjectClass.IsDerivedFrom(RXObject.GetClass(TypeOf (Table))) Then
'            If selId.ObjectClass.IsDerivedFrom(RXObject.GetClass(GetType(Table))) Then

'                tbId = selId

'            End If
'        End If


'        If tbId = ObjectId.Null Then 'If no Table already selected, ask the user to pick one

'            Dim peo = New PromptEntityOptions("\nSelect a table")

'            peo.SetRejectMessage("\nEntity is not a table.")

'            peo.AddAllowedClass(GetType(Table), false)

'            Dim per = ed.GetEntity(peo)

'            If (per.Status <> PromptStatus.OK)
'                Return
'            End If

'            tbId = per.ObjectId

'        End If

'        ' Ask the user to select a destination CSV file 

'        Dim psfo = New PromptSaveFileOptions("Export Data")
'        psfo.Filter = "Comma Delimited (*.csv)|*.csv"
'        Dim pr = ed.GetFileNameForSave(psfo)

'        If (pr.Status <> PromptStatus.OK) Then
'            Return        
'        End If

'        Dim csv = pr.StringResult

'        ' Our main StringBuilder to get the overall CSV contents
'        'Dim sb = New StringBuilder()
'        Dim sb = New System.Text.StringBuilder()

'        Using tr = db.TransactionManager.StartTransaction()

'            Dim tb As Table = tr.GetObject(tbId, OpenMode.ForRead)

'            ' Should be a table but we'll check, just in case

'            If tb <> Nothing Then

'                'For (int i = 0; i < tb.Rows.Count; i++)
'                For i = 0 To (tb.Rows.Count - 1)

'                    'For (int j = 0; j < tb.Columns.Count; j++)
'                    For j = 0 To (tb.Columns.Count - 1)

'                        If (j > 0) Then
'                            sb.Append(",")
'                        End if

'                        ' Get the contents of our cell
'                        Dim c = tb.Cells(i, j)

'                        Dim s = c.GetTextString(FormatOption.ForEditing)

'                        ' This StringBuilder Is for the current cell
'                        'Dim sb2 = New StringBuilder()
'                        Dim sb2 = New System.Text.StringBuilder()

'                        '' Create an MText to access the fragments
'                        '             Using mt = New MText()

'                        '  mt.Contents = s
'                        '  Dim fragNum = 0

'                        '  mt.ExplodeFragments(
'                        '              (frag, obj) => (
'                        '      ' We'll put spaces between fragments


'                        '              ' if (fragNum++ > 0)
'                        '      If ((fragNum + 1) > 0) Then
'                        '        sb2.Append(" ")
'                        '      End if



'                        '      ' As well as replacing any control codes



'                        '      sb2.Append( ReplaceControlCodes(frag.Text))
'                        '              End Using



'                        '      'Return MTextFragmentCallbackStatus.Continue
'                        '          )
'                        '          )

'                    Next
'                Next
'            End if
'        End Using



'        ' And we'll escape strings that require it

'        ' before appending the cell to the CSV string



'        sb.Append( Escape(sb2.ToString()))

'              End if
'next



'            'After each row we start a New line



'            sb.AppendLine();
'next

'        End if



'        tr.Commit()

'      End Using



'      // Get the contents we want to put in the CSV file



'      var contents = sb.ToString();

'      If (!String.IsNullOrWhiteSpace(contents))

'      {

'        Try

'        {

'          // Write the contents to the selected CSV file



'          Using (

'            var sw = New StreamWriter(csv, False, Encoding.UTF8)

'          )

'          {

'            sw.WriteLine(sb.ToString());

'          }

'        }

'        Catch (System.IO.IOException)

'        {

'          // We might have an exception, if the CSV Is open in

'          // Excel, for instance... could also show a messagebox



'          ed.WriteMessage("\nUnable to write to file.");

'        }

'      }

'    }



'    Public Static String ReplaceControlCodes(String s)

'    {

'      // Check the string for each of our control codes, both

'      // upper And lowercase



'      For (int i= 0; i < CODES.Length; i++)

'      {

'        var c = "%%" + CODES[i];

'        If (s.Contains(c))

'        {

'          s = s.Replace(c, REPLS[i]);

'        }

'        var c2 = c.ToLower();

'        If (s.Contains(c2))

'        {

'          s = s.Replace(c2, REPLS[i]);

'        }

'      }



'      Return s;

'    }



'    // AutoCAD control codes And their Unicode replacements

'    // (Codes will be prefixed with "%%")



'    Private Static String[] CODES = {"C", "D", "P"};

'    Private Static String[] REPLS = {"\u00D8", "\u00B0", "\u00B1"};



'    Public Static String Escape(String s)

'    {

'      If (s.Contains(QUOTE))

'                                                                s = s.Replace(QUOTE, ESCAPED_QUOTE);



'      If (s.IndexOfAny(MUST_BE_QUOTED) > -1)

'                                                                    s = QUOTE + s + QUOTE;



'      Return s;

'    }



'    // Constants used to escape the CSV fields



'    Private Const String QUOTE = "\"";

'    private const string ESCAPED_QUOTE = " \ "\"";

'    private static char[] MUST_BE_QUOTED = {',', '" ', '\n'};

' End Sub
'End Class

''}