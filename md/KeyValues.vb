Option Explicit

' Class script: KeyValues

Private Sub SL_KeyValues_LocateAlternatives(ByVal pXDoc As CASCADELib.CscXDocument, ByVal pLocator As CASCADELib.CscXDocField)
   'Extract KeyValue Pairs from Microsoft Document Intelligence
   Dim JS As Object, K As Long, Alt As CscXDocFieldAlternatives, Key As CscXDocSubField, Value As CscXDocSubField, W As Long, KVKey As String
   Set JS= JSON_Parse(Cache_Load(pXDoc, "MicrosoftDI_JSON"))
   For K=0 To CLng(JS("js.analyzeResult.keyValuePairs._count"))-1
      With pLocator.Alternatives.Create
         If K = 9 Then
            K=K
         End If
         KVKey="js.analyzeResult.keyValuePairs(" & K & ")"
         .Confidence=CDouble(JS(KVKey & ".confidence"))
         KVKey= KVKey & ".key.boundingRegions(0)"
         Set Key= .SubFields.Create("Key")
         Key.PageIndex=CLng(JS(KVKey & ".pageNumber")-1)
         JSON_Polygon2Rectangle(JS,KVKey,Key)
         Object_AddWords(pXDoc,Key,Key)
         Key.Confidence=.Confidence
         KVKey="js.analyzeResult.keyValuePairs(" & K & ").value.boundingRegions(0)"
         Set Value = .SubFields.Create("Key")
         Value.PageIndex=CLng(JS(KVKey & ".pageNumber")-1)
         If Value.PageIndex>-1 Then ' Microsoft can return keys without values !!!
            JSON_Polygon2Rectangle(JS,KVKey,Value)
            Object_AddWords(pXDoc,Value,Value)
            Value.Confidence=.Confidence
         End If
      End With
   Next
End Sub

Public Sub Object_AddWords(pXDoc As CscXDocument, o As Object, Region As Object)
   'Add the OCR words from the region to the object. Both can be Field, Locator, Alternative, Subfield or Table Cell
   Dim W As Long, Words As CscXDocWords
   Set Words=pXDoc.GetWordsInRect(Region.PageIndex,Region.Left,Region.Top,Region.Width,Region.Height)
   For W=0 To Words.Count-1
      If TypeOf Region Is CscXDocTableCell Then
         Region.AddWordData(Words(W))  'Table cells handle words differently than fields, locs, alts and subfields
      Else
         Region.Words.Append(Words(W))
      End If
   Next
   Region.Text=Words.Text
End Sub
