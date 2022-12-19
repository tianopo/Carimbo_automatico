Attribute VB_Name = "Image_change"
Public Sub Carimbar()
On Error Resume Next

Dim caminhoImagem As String
Dim Img As String
Dim Foto As Shape
Dim Number As Integer


Img = "Image Files PNG (*.png),*.png,Image Files JPG (*.jpg),*.jpg"

Range("B53:C53").Select
caminhoImagem = Application.GetOpenFilename(Img)

If caminhoImagem <> "" And caminhoImagem <> "Falso" Then
    ActiveSheet.Pictures.Insert(caminhoImagem).Select
    
    Foto = Worksheets(1).Shapes.AddPicture _
    (caminhoImagem, LinktoFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=-1, Top:=-1, Width:=0, Height:=0)
    'Atributo para remoção do vínculo
    
    Selection.ShapeRange.Height = 50
    
    Range("B53:D58").Select
    
    Application.CutCopyMode = False
    
    Selection.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    
    ActiveSheet.Paste
    
    Number = 1
    
    Do Until Number > Worksheets.Count
        Worksheets(Number).Protect Password:="De25Mendes!"
        Number = Number + 1
    Loop
    ActiveWorkbook.Save
End If
End Sub
