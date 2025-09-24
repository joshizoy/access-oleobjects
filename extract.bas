Option Compare Database
Option Explicit

Public Sub ExtractOLEObjects()
    Dim rs As DAO.Recordset
    Dim f As Integer
    Dim fileData() As Byte
    Dim filePath As String

    ' Adjust table and field names
    Set rs = CurrentDb.OpenRecordset("SELECT ID, OLEField FROM MyTable WHERE OLEField Is Not Null")

    Do While Not rs.EOF
      fileData = rs!OLEField   ' Raw binary data
      filePath = "C:\ExportedFiles\Doc_" & rs!ID & ".bin"

      f = FreeFile
      Open filePath For Binary As #f
      Put #f, , fileData
      Close #f

      Debug.Print "Extracted to: " & filePath
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub
