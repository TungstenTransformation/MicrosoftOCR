Option Explicit

' Class script: tax.us.w2

Private Sub Document_AfterLocate(ByVal pXDoc As CASCADELib.CscXDocument, ByVal LocatorName As String)
   Select Case LocatorName
   Case "TL_W2_StateTaxInfos", "TL_W2_LocalTaxInfos"
      W2_TaxTable(pXDoc,LocatorName)
   End Select
End Sub

Private Sub W2_TaxTable(ByVal pXDoc As CASCADELib.CscXDocument, ByVal LocatorName As String)
   Dim JSONs As String, JSON As Object, FieldName As String, Table As CscXDocTable, R As Long, JRow As Object, Row As CscXDocTableRow, colName As String
   Dim Word As CscXDocWord,XRes As Long, YRes As Long, Units As String, Cell As CscXDocTableCell, JCell As Object, Fields As Object, Field As Object
   Dim boundingRegions As Object
   If pXDoc.ExtractionClass="" Then Err.Raise(346,,"Please classify the XDocument before running Locator " & LocatorName)
   Set Table=pXDoc.Locators.ItemByName(LocatorName).Alternatives(0).Table
   Table.Rows.Clear 'ignore anything Transformation may have found by accident
   JSONs=Cache_Load(pXDoc,"MicrosoftDI_JSON",False) 'Get the Microsoft DI response JSON if it is there.
   Set JSON=JSON_Parse(JSONs)
   If JSON("analyzeResult")("documents").Count=0 Then Exit Sub ' Microsoft does not think this is a W2 document
   Set Fields=JSON("analyzeResult")("documents")(0)("fields")
   FieldName=Mid(LocatorName,7)
   If Not Fields.Exists(FieldName) Then Exit Sub ' table is not in JSON
   Set Field=Fields(FieldName)
   For R=0 To Field("valueArray").Count-1
      Set Row=Table.Rows.Append
      Set JRow=Field("valueArray")(R)("valueObject")
      For Each colName In JRow
         Set JCell=JRow(colName)
         Set Word = New CscXDocWord
         Set boundingRegions=JCell("boundingRegions")
         Word.PageIndex=boundingRegions(0)("pageNumber")-1
         Units=JSON("analyzeResult")("pages")(0)("unit")
         XRes=pXDoc.CDoc.Pages(Word.PageIndex).XRes
         YRes=pXDoc.CDoc.Pages(Word.PageIndex).XRes
         BoundingBox2Rectangle(boundingRegions(0)("polygon"),Word, Units, XRes, YRes)
         Word.Text=JCell("content")
         Row.Cells.ItemByName(colName).AddWordData(Word)
      Next
   Next
End Sub


Private Sub SL_W2_LocateAlternatives(ByVal pXDoc As CASCADELib.CscXDocument, ByVal pLocator As CASCADELib.CscXDocField)
   Dim JSONs As String, JSON As Object, Fields As Object, LocDef As CscLocatorDef, FieldName As String, AttributeName As String, S As Long, ClassName As String
   Dim Attributes As Object,Att As Object, Word As CscXDocWord, SubField As CscXDocSubField, A As Long, XRes As Long, YRes As Long, Units As String
   Dim Alt As CscXDocFieldAlternative, Field As Object
   Dim boundingRegions As Object
   ClassName = pXDoc.ExtractionClass
   If ClassName="" Then ClassName="tax.us.w2"
   JSONs=Cache_Load(pXDoc,"MicrosoftDI_JSON",False) 'Get the Microsoft DI response JSON if it is there.
   Set LocDef = Project.ClassByName(ClassName).Locators.ItemByName(pLocator.Name)
   'create all subfields in locator
   Set Alt= pLocator.Alternatives.Create
   Alt.Confidence=1
   For S=0 To LocDef.SubFieldCount-1
      Alt.SubFields.Create(LocDef.SubFieldName(S))
   Next
   'Read fields from Document Intelligence
   Set JSON=JSON_Parse(JSONs)
   If JSON("analyzeResult")("documents").Count=0 Then Exit Sub ' Microsoft does not think this is a W2 document
   Set Fields=JSON("analyzeResult")("documents")(0)("fields")
   For Each FieldName In Fields.keys
      Set Field=Fields(FieldName)
      If Field("type")="object" Then 'Employee contains 3 attributes (SSN, Name, Address)
         Set Attributes=Field("valueObject")
         For Each AttributeName In Attributes.keys
            Set Att=Attributes(AttributeName)
            Set Word = New CscXDocWord
            Set boundingRegions=Att("boundingRegions")
            Word.PageIndex=boundingRegions(0)("pageNumber")-1
            Units=JSON("analyzeResult")("pages")(0)("unit")
            XRes=pXDoc.CDoc.Pages(Word.PageIndex).XRes
            YRes=pXDoc.CDoc.Pages(Word.PageIndex).XRes
            BoundingBox2Rectangle(boundingRegions(0)("polygon"),Word, Units, XRes, YRes)
            Word.Text=Att("content")
            With Alt.SubFields(LocDef.SubFieldNameIndex(FieldName & "_" & AttributeName))
               .Words.Append(Word)
               .Confidence=Att("confidence")
            End With
         Next
      ElseIf Field("type")="array" Then 'Employee contains 3 attributes (SSN, Name, Address)
         'skip arrays - they go into a table locator
      Else
         Set Word = New CscXDocWord
         Set boundingRegions=Field("boundingRegions")
         Word.PageIndex=boundingRegions(0)("pageNumber")-1
         Units=JSON("analyzeResult")("pages")(0)("unit")
         XRes=pXDoc.CDoc.Pages(Word.PageIndex).XRes
         YRes=pXDoc.CDoc.Pages(Word.PageIndex).XRes
         BoundingBox2Rectangle(boundingRegions(0)("polygon"),Word, Units, XRes, YRes)
         Word.Text=Field("content")
         With Alt.SubFields(LocDef.SubFieldNameIndex(FieldName))
            .Words.Append(Word)
            .Confidence=Field("confidence")
         End With
      End If
   Next
End Sub
