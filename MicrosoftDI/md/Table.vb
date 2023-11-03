Option Explicit

' Class script: Table

Private Sub Document_AfterLocate(ByVal pXDoc As CASCADELib.CscXDocument, ByVal LocatorName As String)
   Select Case LocatorName
   Case "TL"
      ATL_Microsoft(pXDoc,LocatorName)
   End Select
End Sub

Public Sub ATL_Microsoft(ByVal pXDoc As CASCADELib.CscXDocument, ByVal LocatorName As String)
      Dim JS As Object, Table As CscXDocTable
      Set Table= pXDoc.Locators.ItemByName(LocatorName).Alternatives(0).Table
      Set JS= JSON_Parse(Cache_Load(pXDoc, "MicrosoftDI_JSON"))
      MicrosoftDI_AddTable(pXDoc,JS,pXDoc.Locators.ItemByName(LocatorName).Alternatives(0).Table,0)
End Sub
