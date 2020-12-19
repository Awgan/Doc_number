Attribute VB_Name = "ThisWorkbook_"
Public Sub workbook_activate()

Dim n As Integer

Dim sh As Worksheet
Set sh = Worksheets(1)

'Generowanie numeru dnia bierz¹cego roku
n = DatePart("y", Date)
sh.Range("A1").Value = n

ActiveWorkbook.Save

End Sub


