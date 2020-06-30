Attribute VB_Name = "Módulo1"
Public Sub getlabels()
    
    Application.ScreenUpdating = False
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFiles As Object
    Dim rng As Range
    Dim ws2 As Worksheet
    Dim lLastCol As Long
    Dim lLastRow As Long
    Dim Lista() As String
    
    'Definindo a pasta base para puxar arquivos
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(Application.ThisWorkbook.Path & "\base")
    Set oFiles = oFolder.Files
    
    'Adicionando novas Sheets no baselabels de acordo com o número de arquivos na pasta
    
    cont = 0
    qtd_arquivos = oFiles.Count - 1
    
    If qtd_arquivos <> 0 Then
        ActiveWorkbook.Sheets.Add Count:=qtd_arquivos
    End If
    
    'Iterando sobre os arquivos da pasta para extrair a baselabel
    
    For Each file In oFiles
        cont = cont + 1
        
        Workbooks.Open (oFolder & "\" & file.Name)
        Set ws2 = Workbooks(file.Name).Sheets(1)
        
    'Definindos Campos do arquivo aberto que devem ser extraídos 
        Application.ScreenUpdating = True
        
        tam_array = Application.InputBox("Enter the Number of Fields from target to be extracted", Type:=1)
    
        ReDim Lista(tam_array)
    
        For j = 0 To tam_array - 1
            Lista(j) = Application.InputBox("Enter the Fields Names from target Table to be Extracted Labels", Type:=2)
        Next j
        
        Application.ScreenUpdating = False
        
    'Iterando
        
        ThisWorkbook.Sheets(cont).Name = file.Name
        
        lColumn = ws2.Cells(1, Columns.Count).End(xlToLeft).Column
        
        cont2 = 1
        
        For i = 1 To lColumn
            'If Cells(1, i).Text = "City" Or Cells(1, i).Text = "State" Then
            If IsInArray(Cells(1, i).Text, Lista) Then
                ws2.Columns(i).RemoveDuplicates Columns:=1, Header:=xlNo
                lLastRow = ws2.Cells(Rows.Count, i).End(xlUp).Row
                Set rng = ws2.Range(Cells(1, i), Cells(lLastRow, i))
                rng.Copy Destination:=ThisWorkbook.Worksheets(file.Name).Columns(cont2)
                cont2 = cont2 + 1
            End If
        Next i
        Workbooks(file.Name).Close SaveChanges:=False
    Next file
    
    Application.ScreenUpdating = True
End Sub
Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function
