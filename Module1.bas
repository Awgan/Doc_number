Attribute VB_Name = "Module1"
Sub przypisz_numer()

Dim lp, n As Integer
Dim nr, opis As String
Dim znak_roku As String

znak_roku = "ZA"


'Generowanie numeru dnia bierz¹cego roku
n = DatePart("y", Date)
Range("A1").Value = n



If Range("A1").Value = Range("A2").Value Then
      
      
  lp = Range("B1").Value + 1
         
         
            If n >= 100 Then
                                        If lp < 10 Then
                                        
                                        'MsgBox ("T" & n & "0" & lp & "RS")
                                        Range("A2").Value = DatePart("y", Date)
                                        nr = znak_roku & n & "0" & lp & "RS"
                                        
                                                    
                                        Else
                                        
                                        'MsgBox ("T" & n & lp & "RS")
                                        Range("A2").Value = DatePart("y", Date)
                                        nr = znak_roku & n & lp & "RS"
                                        End If
        
        Else
        
                                        If lp < 10 Then
                                        
                                        'MsgBox ("T" & n & "0" & lp & "RS")
                                        Range("A2").Value = DatePart("y", Date)
                                        nr = znak_roku & "0" & n & "0" & lp & "RS"
                                        
                                                    
                                        Else
                                        
                                        'MsgBox ("T" & n & lp & "RS")
                                        Range("A2").Value = DatePart("y", Date)
                                        nr = znak_roku & "0" & n & lp & "RS"
                                        
                                        End If
        
        End If
        
        
        Range("B1").Value = lp

Else

                                        If n >= 100 Then
                                        
                                                    Range("B1").Value = ""
                                                    Range("C:C").Value = ""
                                                    lp = Range("B1").Value + 1
                                                    'MsgBox ("T" & n & "0" & lp & "RS")
                                                    nr = znak_roku & n & "0" & lp & "RS"
                                                    Range("A2").Value = DatePart("y", Date)
                                                    Range("B1").Value = lp
                                        
                                        Else
                                        
                                                    Range("B1").Value = ""
                                                    Range("C:C").Value = ""
                                                    lp = Range("B1").Value + 1
                                                    'MsgBox ("T" & n & "0" & lp & "RS")
                                                    nr = znak_roku & "0" & n & "0" & lp & "RS"
                                                    Range("A2").Value = DatePart("y", Date)
                                                    Range("B1").Value = lp
                                        End If



End If

    
    
    
            Worksheets(2).Select
            Range("A1").Select
              
            Dim LastRow As Long
                
                With ActiveSheet
                    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
                End With
            'MsgBox LastRow
             
            ActiveCell.Offset(LastRow, 0).Select
            ActiveCell.Value = nr

    opis = InputBox("Opis dokumentu nr:" & nr)
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = opis
    
    Worksheets(1).Select
    Range("C1").Value = nr
    Range("C1").Copy
    
ActiveWorkbook.Save


End Sub
