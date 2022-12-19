Attribute VB_Name = "Módulo1"
Sub Limpar_Planilha_de_Pedidos()
Attribute Limpar_Planilha_de_Pedidos.VB_Description = "Esta macro limpa os campos para iniciar um novo pedido"
Attribute Limpar_Planilha_de_Pedidos.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Limpar_Planilha_de_Pedidos Macro
' Esta macro limpa os campos para iniciar um novo pedido
'
'
    Range("B8:G8").Select
    Selection.ClearContents
    Range("B11:G11").Select
    Selection.ClearContents
    Range("B12:D12").Select
    Selection.ClearContents
    Range("B13:D13").Select
    Selection.ClearContents
    Range("B14:D14").Select
    Selection.ClearContents
    Range("B15:D15").Select
    Selection.ClearContents
    Range("F12:G12").Select
    Selection.ClearContents
    Range("F13:G13").Select
    Selection.ClearContents
    Range("F14:G14").Select
    Selection.ClearContents
    Range("F15:G15").Select
    Selection.ClearContents
    Range("C18:G18").Select
    Selection.ClearContents
    Range("C19:G19").Select
    Selection.ClearContents
    Range("C20:G20").Select
    Selection.ClearContents
    Range("C21:G21").Select
    Selection.ClearContents
    Range("A24:F41").Select
    Selection.ClearContents
    Range("A1:G1").Select
End Sub
Sub Salvar_Pedido_em_PDF()
Attribute Salvar_Pedido_em_PDF.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Salvar_Pedido_em_PDF Macro
'

'

CaminhoPersonalizado = Left(ThisWorkbook.FullName, Len(ThisWorkbook.FullName) - Len(ThisWorkbook.Name))
Nome = "Pedido_" & Sheets("Planilha1").Range("f4").Value

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        CaminhoPersonalizado & Nome & ".pdf", _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
        True
End Sub
