Option Explicit

' Class script: Table

Private Sub Document_AfterLocate(ByVal pXDoc As CASCADELib.CscXDocument, ByVal LocatorName As String)
   Select Case LocatorName
   Case "ATL"
      ATL_Microsoft(pXDoc,LocatorName)
   End Select
End Sub

Public Sub ATL_Microsoft(ByVal pXDoc As CASCADELib.CscXDocument, ByVal LocatorName As String)
      Dim JS As Object, FileName As String, Table As CscXDocTable
      Set Table= pXDoc.Locators.ItemByName(LocatorName).Alternatives(0).Table
      FileName=Replace(pXDoc.FileName,".xdc","." & Project.ScriptVariables("MicrosoftFormRecognizerModel") &".json")
      Set JS= JSON_Parse(File_Load(FileName))
      MicrosoftOCR_AddTable(pXDoc,JS,pXDoc.Locators.ItemByName(LocatorName).Alternatives(0).Table,0)
End Sub
