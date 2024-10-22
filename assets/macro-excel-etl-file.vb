Option Explicit

Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim lLinha As Long

'Função que identifica a existência do arquivo
Private Function lfVerificaArquivo(ByVal lStr As String) As Boolean

    lfVerificaArquivo = True
    
    'Identifica se o arquivo existe
    If Dir(lStr) = vbNullString Then
        lfVerificaArquivo = False
'       Worksheets("BASELINE ENG M").Cells(lLinha, 24).Value = "Indisponível"
        Worksheets("PROJETO").Cells(lLinha, 12).Value = Empty
    Else
'        Worksheets("PROJETO").Cells(lLinha, 24).Value = "OK"
        Worksheets("PROJETO").Cells(lLinha, 12).FormulaR1C1 = "=IF(LEFT(RC[-6],1)=""C"","""",HYPERLINK(""\\192.175.175.4\dados\EMPRESA\PROJETOS\PDF-OFICIAL\""&RC[-6]&"".png"",""PNG""))"
    End If
    
End Function

'Procedimento que realiza um loop por todos os arquivos de configuração
Public Sub PNG()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    lib_filtro_projeto
    
    Worksheets("PROJETO").Range("L2:L10000").ClearContents
    
    Dim lUltimaLinhaAtiva   As Long
    Dim endereço As String
    
    lLinha = 2
    
    'Identifica a quantidade de linhas preenchidas
    lUltimaLinhaAtiva = Worksheets("PROJETO").Cells(Rows.Count, 5).End(xlUp).Row
    
    'Realiza um loop por todos os registros
    While lLinha <= lUltimaLinhaAtiva
    
    endereço = "\\192.175.175.4\dados\EMPRESA\PROJETOS\PDF-OFICIAL\" & Worksheets("PROJETO").Cells(lLinha, 6).Value & ".png"
        'Se não for encontrado um arquivo o procedimento é abortado
        If lfVerificaArquivo(endereço) = False Then
            'Exit Sub
        End If
        lLinha = lLinha + 1
    Wend
    
    MsgBox "Mapeamento de arquivos PDF concluído"
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    
    lUltimaLinhaAtiva = Empty
    endereço = Empty
    lLinha = Empty
End Sub

Sub GERAR_BASELINE()

Dim uln, uln1 As Long
Dim i, x As Integer
Dim xList As ListObject

Application.ScreenUpdating = False

Sheets("PROJETO").Select

For Each xList In Sheets("PROJETO").ListObjects
    xList.Unlist
Next

Sheets("BASELINE").Select

uln = Worksheets("PROJETO").Cells(Rows.Count, 5).End(xlUp).Row + 1

If Sheets("BASELINE").Range("A9").Value > 0 Or Sheets("BASELINE").Range("A12").Value > 1 Then
    If Sheets("BASELINE").Range("A12").Value > uln Then
        MsgBox "LINHA DEVE SER MENOR OU IGUAL A ULTIMA LINHA DA PLANILHA PROJETO"
        End
    Else
        If Sheets("BASELINE").Range("A12").Value < 1 Or Sheets("BASELINE").Range("A12").Value = "" Then
            MsgBox "NIVEIS MAIOR QUE 0 É NECESSÁRIO DEFINIR A LINHA DE INSERÇÃO"
            End
        End If
    End If
End If

uln = Empty

Call lib_filtro_projeto

Sheets("BASELINE").Select
    
uln = Worksheets("BASELINE").Cells(Rows.Count, 13).End(xlUp).Row
    
Columns("N:S").ClearContents

Sheets("BASELINE").Range("U2:BB100000").ClearContents

Application.Calculation = xlManual

Call titulos

    Range("N2").FormulaR1C1 = "=SUBSTITUTE(TEXTJOIN(""|"",TRUE,RC[-11]:RC[-1]),""."","","")"
    Range("O2").FormulaR1C1 = "=RC[-1]"
    Range("O3").FormulaR1C1 = "=MID(RC[-1],1,SEARCH(""|"",RC[-1])-1)"
    Range("P2").Formula2R1C1 = "=IF(RC[-1]=""(vazio)"",""(vazio)"",XLOOKUP(RC[1]-1,R[-1]C17:R2C17,R[-1]C[-1]:R2C15,"""",0,-1))"
    Range("Q2").FormulaR1C1 = "=0+R[7]C[-16]"
    Range("Q3").FormulaR1C1 = "=MATCH(MID(RC[-3],1,SEARCH(""|"",RC[-3])-1),RC3:RC13,0)+R9C1"
    Range("R2").FormulaR1C1 = "=IF(RC[-3]=""(vazio)"","""",CONCAT(REPT(""."",RC[-1]),RC[-3]))"
    Range("S2").FormulaR1C1 = "1"
    Range("S3").FormulaR1C1 = "=IF(RC[-1]="""","""",VALUE(MID(RC[-5],SEARCH(""|"",RC[-5])+1,10)))"
    Range("Y2").Formula2R1C1 = "=FILTER(R2C17:R54337C17,R2C18:R54337C18<>"""")"
    
'    Range("Z2").Formula2R1C1 = "=TRIM(FILTER(R2C18:R54337C18,R2C18:R54337C18<>""""))"
    Range("Y2").Formula2R1C1 = "=FILTER(R2C17:R54337C17,R2C18:R54337C18<>"""")"
    Range("Z2").Formula2R1C1 = "=TRIM(FILTER(R2C15:R54337C15,R2C18:R54337C18<>""""))"
    Range("AA2").Formula2R1C1 = "=TRIM(FILTER(R2C16:R54337C16,R2C18:R54337C18<>""""))"
    Range("AD2").Formula2R1C1 = "=FILTER(R2C19:R54337C19,R2C18:R54337C18<>"""")"
    
    Range("N2").Select
    
    Cells(2, 14).Select
    Selection.AutoFill Destination:=Range(Cells(2, 14), Cells(uln, 14))
    
    Cells(3, 15).Select
    Selection.AutoFill Destination:=Range(Cells(3, 15), Cells(uln, 15))
    
    Cells(2, 16).Select
    Selection.AutoFill Destination:=Range(Cells(2, 16), Cells(uln, 16))
    
    Cells(3, 17).Select
    Selection.AutoFill Destination:=Range(Cells(3, 17), Cells(uln, 17))
    
    Cells(2, 18).Select
    Selection.AutoFill Destination:=Range(Cells(2, 18), Cells(uln, 18))
    
    Cells(3, 19).Select
    Selection.AutoFill Destination:=Range(Cells(3, 19), Cells(uln, 19))
    
    Application.Calculation = xlAutomatic
    
    Range(Cells(2, 14), Cells(uln, 19)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("A2").Select

uln = Empty

uln = Worksheets("BASELINE").Cells(Rows.Count, 25).End(xlUp).Row

    Application.Calculation = xlManual
    
    Range(Cells(2, 35), Cells(uln, 35)).Select
    Selection.FormulaR1C1 = "TEMPORÁRIO"

i = Empty

Range(Cells(2, 25), Cells(uln, 35)).Copy
Range("Y2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
Application.CutCopyMode = False


uln1 = Worksheets("PROJETO").Cells(Rows.Count, 5).End(xlUp).Row + 1

Dim Number

If Sheets("BASELINE").Range("A12").Value <> "" Then
    Number = 1
Else
    Number = 0
End If

Range(Cells(2, 21), Cells(uln, 52)).Select

uln = Empty

Selection.Copy

Select Case Number

Case 0
    Worksheets("PROJETO").Select
    Worksheets("PROJETO").Cells(uln1, 1).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
    Application.CutCopyMode = False
Case 1
    x = Worksheets("BASELINE").Cells(12, 1).Value
    Worksheets("PROJETO").Select
    Worksheets("PROJETO").Cells(x, 1).Select
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
Case Else

End Select

Worksheets("BASELINE").Select

Worksheets("BASELINE").Range("N:BA").ClearContents

Worksheets("PROJETO").Select

Call ID

uln = Empty
i = Empty
Number = Empty
x = Empty

Sheets("BASELINE").Range("A9").Value = Empty
Sheets("BASELINE").Range("A12").Value = Empty

Call TAB_PROJETO
Call DESCRIÇÕES

Cells.Select
Cells.EntireColumn.AutoFit
Range("A1").Select

'Call definir_tabela
Call FORMT_CONDICIONAL

Application.Calculation = xlAutomatic
Application.ScreenUpdating = True

End Sub

Sub CONEXAO_TOTVs()
Range("P1").Value = Now

    Dim tb As ListObject
    Set tb = Planilha25.ListObjects("TAB_PRODUTO")
    tb.Refresh


On Error Resume Next
ActiveSheet.PivotTables("BASELINE").PivotCache.Refresh
Range("P2").Value = Now
End Sub

Sub REPLICAR()

Application.ScreenUpdating = False

    Dim loc As String
    loc = ActiveCell.Address

lib_filtro_projeto

Dim nivel, i As Long

If ActiveCell.Column = 15 Then
    nivel = ActiveCell.Offset(0, -10).Value
    i = 1
    While ActiveCell.Offset(i, -10).Value > nivel
        If ActiveCell.Offset(i, 0).Value = "" Then
            ActiveCell.Offset(i, 0).Value = ActiveCell.Value
        End If
        i = i + 1
    Wend
End If

If ActiveCell.Column = 16 Then
    nivel = ActiveCell.Offset(0, -11).Value
    i = 1
    While ActiveCell.Offset(i, -11).Value > nivel
        If ActiveCell.Offset(i, 0).Value = "" Then
            ActiveCell.Offset(i, 0).Value = ActiveCell.Value
        End If
        i = i + 1
    Wend
End If

If ActiveCell.Column = 17 Then
    nivel = ActiveCell.Offset(0, -12).Value
    i = 1
    While ActiveCell.Offset(i, -12).Value > nivel
        If ActiveCell.Offset(i, 0).Value = "" Then
            ActiveCell.Offset(i, 0).Value = ActiveCell.Value
        End If
        i = i + 1
    Wend
End If

Range(loc).Select

Application.ScreenUpdating = True

nivel = Empty
i = Empty

End Sub

Sub CHANGE()

Application.ScreenUpdating = False

lib_filtro_projeto

If ActiveCell.Row > 1 And ActiveCell.Column = 6 And ActiveCell.Value <> "" Then

Dim i As Long

i = ActiveCell.Row

Cells(i + 1, 6).Select
    Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
    
    Call inserir
    Call ID

Cells(i + 1, 6).Select

End If

Application.ScreenUpdating = True

i = Empty
    
End Sub
Function ID()

Range("A:A").ClearContents
    Range("A1").Value = "ID"
    Range("A2").Select
    ActiveCell.Formula2R1C1 = "=SEQUENCE(COUNTA(C[4])-1,1,1)"
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select
    Application.CutCopyMode = False

End Function
Function inserir()

Dim i As Long
i = ActiveCell.Row

    Rows(i - 1).Select
    Selection.Copy
    Rows(i).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Cells(i, 7).ClearContents
    Cells(i, 8).ClearContents
    Cells(i, 9).ClearContents
    Cells(i, 10).ClearContents
    Cells(i, 11).ClearContents
    Cells(i, 12).ClearContents
    Cells(i, 13).ClearContents
    Cells(i, 14).ClearContents
    Cells(i, 15).ClearContents
    Cells(i, 16).ClearContents
    
    Cells(i - 1, 15).Value = "SUBSTITUÍDO"
    Cells(i - 1, 18).Value = 0
    
i = Empty
    
End Function

Sub QUANTIDADE()

Application.ScreenUpdating = False

    Dim loc As String
    loc = ActiveCell.Address

Call lib_filtro_projeto

Dim uln, uln1 As Long

Sheets("PROJETO").Select

uln = Worksheets("PROJETO").Cells(Rows.Count, 5).End(xlUp).Row
uln1 = Worksheets("PROJETO").Cells(Rows.Count, 5).End(xlUp).Row

If uln1 > 1 Then
    
Range(Cells(2, 19), Cells(uln, 31)).ClearContents

Application.Calculation = xlManual

    Range("T2").FormulaR1C1 = "=IF(VALUE(RC5)>VALUE(R1C),R[-1]C,IF(VALUE(RC5)=VALUE(R1C),IF((RC17<>""OK""),0,RC18),1))"
    Range("U2").FormulaR1C1 = "=IF(VALUE(RC5)>VALUE(R1C),R[-1]C,IF(VALUE(RC5)=VALUE(R1C),IF((RC17<>""OK""),0,RC18),1))"
    Range("V2").FormulaR1C1 = "=IF(VALUE(RC5)>VALUE(R1C),R[-1]C,IF(VALUE(RC5)=VALUE(R1C),IF((RC17<>""OK""),0,RC18),1))"
    Range("W2").FormulaR1C1 = "=IF(VALUE(RC5)>VALUE(R1C),R[-1]C,IF(VALUE(RC5)=VALUE(R1C),IF((RC17<>""OK""),0,RC18),1))"
    Range("X2").FormulaR1C1 = "=IF(VALUE(RC5)>VALUE(R1C),R[-1]C,IF(VALUE(RC5)=VALUE(R1C),IF((RC17<>""OK""),0,RC18),1))"
    Range("Y2").FormulaR1C1 = "=IF(VALUE(RC5)>VALUE(R1C),R[-1]C,IF(VALUE(RC5)=VALUE(R1C),IF((RC17<>""OK""),0,RC18),1))"
    Range("Z2").FormulaR1C1 = "=IF(VALUE(RC5)>VALUE(R1C),R[-1]C,IF(VALUE(RC5)=VALUE(R1C),IF((RC17<>""OK""),0,RC18),1))"
    Range("AA2").FormulaR1C1 = "=IF(VALUE(RC5)>VALUE(R1C),R[-1]C,IF(VALUE(RC5)=VALUE(R1C),IF((RC17<>""OK""),0,RC18),1))"
    Range("AB2").FormulaR1C1 = "=IF(VALUE(RC5)>VALUE(R1C),R[-1]C,IF(VALUE(RC5)=VALUE(R1C),IF((RC17<>""OK""),0,RC18),1))"
    Range("AC2").FormulaR1C1 = "=IF(VALUE(RC5)>VALUE(R1C),R[-1]C,IF(VALUE(RC5)=VALUE(R1C),IF((RC17<>""OK""),0,RC18),1))"
    Range("AD2").FormulaR1C1 = "=PRODUCT(RC[-10]:RC[-1])"
    Range("AE2").FormulaR1C1 = "=IF(RC[-16]=""SUBSTITUÍDO"",""FINALIZADO"",IF(RC[-16]=""DESCONSIDERAR"",""FINALIZADO"",IF(OR(RC[-16]<>""PRONTO"",RC[-13]="""",RC[-14]<>""OK"",RC[-15]<>""OK""),""PENDENTE"",""FINALIZADO"")))"
    Range("S2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-1]/RC[-9],"""")"
    
Application.Calculation = xlAutomatic
    
Range("S2:AE2").Select

Selection.AutoFill Destination:=Range(Cells(2, 19), Cells(uln, 31))
    
Range(Cells(2, 19), Cells(uln, 31)).Select

Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
    
End If

uln = Empty
uln1 = Empty

Call TRATAMENTO_PI
    
Call TAB_PROJETO

Call FORMT_CONDICIONAL

Range(loc).Select

Application.ScreenUpdating = True

Call ETLSalvarBaselineNoMSSQL
    
End Sub
Sub ETLSalvarBaselineNoMSSQL()
    
    Dim fileName As String
    Dim numeroQP As String
    Dim bomPath As String
    Dim wb As Workbook

    ' Obtem o nome do arquivo Excel que estiver aberto
    Set wb = ActiveWorkbook
    fileName = wb.Name ' Pega apenas o nome do arquivo, sem o caminho
    
    ' Extrai o QP-EXXXX do nome do arquivo usando a função Mid e InStr
    numeroQP = Mid(fileName, InStr(fileName, "QP-"), 8) ' Extrai "QP-EXXXX"
    
    Call DefinirVariavelAmbiente(numeroQP)

    ' Define o caminho do arquivo na pasta TEMP
    bomPath = Environ("TEMP") & "\" & numeroQP & ".xlsm" ' Salvar como .xlsm

    ' Remove o arquivo existente se já estiver presente
    DeleteFileIfExists bomPath
    
    ' Salva uma cópia do arquivo Excel na pasta TEMP como .xlsm
    wb.SaveCopyAs fileName:=bomPath ' Salva cópia como .xlsm, sem fechar o original

    ' Aguarda até que o arquivo seja salvo
    While Dir(bomPath) = ""
    Wend

    ' Chama o script Python
    Call ExecutarScriptPython

End Sub
Sub DefinirVariavelAmbiente(valorVariavel As String)

    ' Substitua "meuarquivo" pelo nome real do seu arquivo (sem a extensão)
    Dim numeroQP As String
    numeroQP = valorVariavel

    ' Construa o comando para definir a variável de ambiente
    Dim comando As String
    comando = "setx QP_BASELINE " & numeroQP

    ' Execute o comando no shell
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run comando, 0, True
    
End Sub
Sub ExecutarScriptPython()
    Dim CaminhoArquivo As String
    
    ' MsgBox "Executando script Python...", vbInformation, "EUREKA®"
    
    ' Substitua o caminho abaixo pelo caminho do seu arquivo botao-salvar-bom-solidworks-totvs
    CaminhoArquivo = "\\192.175.175.4\desenvolvimento\REPOSITORIOS\etl-insert-baseline-mssql-database\etl_baseline_mssql.pyw"
    
    ' Use a funcao Shell para abrir o arquivo com a aplicacao padrao
    Shell "explorer.exe """ & CaminhoArquivo & """", vbNormalFocus
End Sub
Sub DeleteFileIfExists(filePath As String)
    If Dir(filePath) <> "" Then
        ' File exists, so delete it
        Kill filePath
        While Dir(filePath) <> ""
        Wend
    Else
        ' File does not exist
        ' MsgBox "File does not exist: " & filePath
    End If
End Sub

Public Function printthisdoc(formname As Long, fileName As String)
On Error Resume Next
Dim x As Long
x = ShellExecute(formname, "Print", fileName, 0&, 0&, 3)
End Function
End Function

Sub TRATAMENTO_PI()

Application.Calculation = xlManual

    Columns("AF:BM").Select
    Selection.ClearContents
    Range("AF1").FormulaR1C1 = "STATUS_OP"
    Range("AF2").Select

Dim uln As Long

uln = Worksheets("PROJETO").Cells(Rows.Count, 5).End(xlUp).Row

Range("BD1").Formula2R1C1 = "=SEQUENCE(1,10,0)"

Range("AH2").FormulaR1C1 = "=IF(RC[1]>0,COUNTIFS(R1C5:RC5,R1C[22]),IF(RC31=""FINALIZADO"",0,IF(R1C[22]<>RC5,0,COUNTIFS(R1C5:RC5,R1C[22]))))"
Range("AI2").FormulaR1C1 = "=IF(RC[1]>0,COUNTIFS(R1C5:RC5,R1C[22]),IF(RC31=""FINALIZADO"",0,IF(R1C[22]<>RC5,0,COUNTIFS(R1C5:RC5,R1C[22]))))"
Range("AJ2").FormulaR1C1 = "=IF(RC[1]>0,COUNTIFS(R1C5:RC5,R1C[22]),IF(RC31=""FINALIZADO"",0,IF(R1C[22]<>RC5,0,COUNTIFS(R1C5:RC5,R1C[22]))))"
Range("AK2").FormulaR1C1 = "=IF(RC[1]>0,COUNTIFS(R1C5:RC5,R1C[22]),IF(RC31=""FINALIZADO"",0,IF(R1C[22]<>RC5,0,COUNTIFS(R1C5:RC5,R1C[22]))))"
Range("AL2").FormulaR1C1 = "=IF(RC[1]>0,COUNTIFS(R1C5:RC5,R1C[22]),IF(RC31=""FINALIZADO"",0,IF(R1C[22]<>RC5,0,COUNTIFS(R1C5:RC5,R1C[22]))))"
Range("AM2").FormulaR1C1 = "=IF(RC[1]>0,COUNTIFS(R1C5:RC5,R1C[22]),IF(RC31=""FINALIZADO"",0,IF(R1C[22]<>RC5,0,COUNTIFS(R1C5:RC5,R1C[22]))))"
Range("AN2").FormulaR1C1 = "=IF(RC[1]>0,COUNTIFS(R1C5:RC5,R1C[22]),IF(RC31=""FINALIZADO"",0,IF(R1C[22]<>RC5,0,COUNTIFS(R1C5:RC5,R1C[22]))))"
Range("AO2").FormulaR1C1 = "=IF(RC[1]>0,COUNTIFS(R1C5:RC5,R1C[22]),IF(RC31=""FINALIZADO"",0,IF(R1C[22]<>RC5,0,COUNTIFS(R1C5:RC5,R1C[22]))))"
Range("AP2").FormulaR1C1 = "=IF(RC[1]>0,COUNTIFS(R1C5:RC5,R1C[22]),IF(RC31=""FINALIZADO"",0,IF(R1C[22]<>RC5,0,COUNTIFS(R1C5:RC5,R1C[22]))))"
Range("AQ2").FormulaR1C1 = "=IF(RC[1]>0,COUNTIFS(R1C5:RC5,R1C[22]),IF(RC31=""FINALIZADO"",0,IF(R1C[22]<>RC5,0,COUNTIFS(R1C5:RC5,R1C[22]))))"

Range(Cells(2, 34), Cells(2, 43)).Select
Selection.AutoFill Destination:=Range(Cells(2, 34), Cells(uln, 43))

Range("AS2").Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(C[-11],(C31=""PENDENTE"")*(C[-11]>0))),"""")"
Range("AT2").Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(C[-11],(C31=""PENDENTE"")*(C[-11]>0))),"""")"
Range("AU2").Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(C[-11],(C31=""PENDENTE"")*(C[-11]>0))),"""")"
Range("AV2").Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(C[-11],(C31=""PENDENTE"")*(C[-11]>0))),"""")"
Range("AW2").Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(C[-11],(C31=""PENDENTE"")*(C[-11]>0))),"""")"
Range("AX2").Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(C[-11],(C31=""PENDENTE"")*(C[-11]>0))),"""")"
Range("AY2").Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(C[-11],(C31=""PENDENTE"")*(C[-11]>0))),"""")"
Range("AZ2").Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(C[-11],(C31=""PENDENTE"")*(C[-11]>0))),"""")"
Range("BA2").Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(C[-11],(C31=""PENDENTE"")*(C[-11]>0))),"""")"
Range("BB2").Formula2R1C1 = "=IFERROR(UNIQUE(FILTER(C[-11],(C31=""PENDENTE"")*(C[-11]>0))),"""")"

Range("BD2").FormulaR1C1 = "=IF(COUNTIF(C[-11],COUNTIFS(R1C5:RC5,R1C))>0,""x"",R1C)"
Range("BE2").FormulaR1C1 = "=IF(COUNTIF(C[-11],COUNTIFS(R1C5:RC5,R1C))>0,""x"",R1C)"
Range("BF2").FormulaR1C1 = "=IF(COUNTIF(C[-11],COUNTIFS(R1C5:RC5,R1C))>0,""x"",R1C)"
Range("BG2").FormulaR1C1 = "=IF(COUNTIF(C[-11],COUNTIFS(R1C5:RC5,R1C))>0,""x"",R1C)"
Range("BH2").FormulaR1C1 = "=IF(COUNTIF(C[-11],COUNTIFS(R1C5:RC5,R1C))>0,""x"",R1C)"
Range("BI2").FormulaR1C1 = "=IF(COUNTIF(C[-11],COUNTIFS(R1C5:RC5,R1C))>0,""x"",R1C)"
Range("BJ2").FormulaR1C1 = "=IF(COUNTIF(C[-11],COUNTIFS(R1C5:RC5,R1C))>0,""x"",R1C)"
Range("BK2").FormulaR1C1 = "=IF(COUNTIF(C[-11],COUNTIFS(R1C5:RC5,R1C))>0,""x"",R1C)"
Range("BL2").FormulaR1C1 = "=IF(COUNTIF(C[-11],COUNTIFS(R1C5:RC5,R1C))>0,""x"",R1C)"
Range("BM2").FormulaR1C1 = "=IF(COUNTIF(C[-11],COUNTIFS(R1C5:RC5,R1C))>0,""x"",R1C)"

Range(Cells(2, 56), Cells(2, 65)).Select
Selection.AutoFill Destination:=Range(Cells(2, 56), Cells(uln, 65))

Range("AF2").FormulaR1C1 = "=IF(AND(RC[-17]=""PRONTO"",RC[-23]<>""MP"",RC[-2]<>0),IF(SMALL(RC[24]:RC[33],1)=RC[-27],""ABRIR OP"",IF(RC[-27]<SMALL(RC[24]:RC[33],1),"""",""CONTIDO"")),"""")"

Range("AF2").Select
Selection.AutoFill Destination:=Range(Cells(2, 32), Cells(uln, 32))

Calculate
    
Range("AF:BM").Select

Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

Columns("AG:BM").Select
Selection.ClearContents

Range("AF1").Select

Application.Calculation = xlAutomatic

uln = Empty

End Sub

Sub lib_filtro_projeto()
Dim xList As ListObject


Sheets("PROJETO").Select

For Each xList In Sheets("PROJETO").ListObjects
    xList.Unlist
Next

    If Sheets("PROJETO").AutoFilterMode = True Then
    Sheets("PROJETO").AutoFilter.ShowAllData
    End If
    
End Sub

Sub definir_tabela()
Dim uln As Long
uln = Worksheets("PROJETO").Cells(Rows.Count, 5).End(xlUp).Row

Sheets("PROJETO").ListObjects.Add(xlSrcRange, Range("$A$1:$AF$" & uln), , xlYes).Name = "TAB_BASELINE"

ActiveSheet.ListObjects("TAB_BASELINE").TableStyle = "TableStyleMedium4"

End Sub

Sub titulos()

Cells(1, 21).FormulaR1C1 = "ID"
Cells(1, 22).FormulaR1C1 = "VISÃOGERAL"
Cells(1, 23).FormulaR1C1 = "EQUIPAMENTO"
Cells(1, 24).FormulaR1C1 = "GRUPO"
Cells(1, 25).FormulaR1C1 = "NIVEL"
Cells(1, 26).FormulaR1C1 = "CÓDIGO"
Cells(1, 27).FormulaR1C1 = "CÓDIGOPAI"
Cells(1, 28).FormulaR1C1 = "DESCRIÇÃO"
Cells(1, 29).FormulaR1C1 = "TIPO"
Cells(1, 30).FormulaR1C1 = "QTDEBL"
Cells(1, 31).FormulaR1C1 = "UND"
Cells(1, 32).FormulaR1C1 = "LINK"
Cells(1, 33).FormulaR1C1 = "OBSERVAÇÕES"
Cells(1, 34).FormulaR1C1 = "PEÇA REPOSIÇÃO"
Cells(1, 35).FormulaR1C1 = "ESPECIFICAÇÕES"
Cells(1, 36).FormulaR1C1 = "TOTVs"
Cells(1, 37).FormulaR1C1 = "QTDE"
Cells(1, 38).FormulaR1C1 = "QTDEPROJ."
Cells(1, 39).FormulaR1C1 = "%"
Cells(1, 40).FormulaR1C1 = "0"
Cells(1, 41).FormulaR1C1 = "1"
Cells(1, 42).FormulaR1C1 = "2"
Cells(1, 43).FormulaR1C1 = "3"
Cells(1, 44).FormulaR1C1 = "4"
Cells(1, 45).FormulaR1C1 = "5"
Cells(1, 46).FormulaR1C1 = "6"
Cells(1, 47).FormulaR1C1 = "7"
Cells(1, 48).FormulaR1C1 = "8"
Cells(1, 49).FormulaR1C1 = "9"
Cells(1, 50).FormulaR1C1 = "QTDETOTAL"
Cells(1, 51).FormulaR1C1 = "STATUS"
Cells(1, 52).FormulaR1C1 = "STATUS_OP"

Cells(2, 1).Select
    
End Sub

Sub DESCRIÇÕES()

Application.ScreenUpdating = False

    Dim loc As String
    loc = ActiveCell.Address

If Sheets("PROJETO").Range("F2").Value <> "" Then

Call lib_filtro_projeto

Call definir_tabela

    Range("TAB_BASELINE[DESCRIÇÃO]").Select
    Selection.ClearContents
    Range("TAB_BASELINE[TIPO]").Select
    Selection.ClearContents
    Range("TAB_BASELINE[UND]").Select
    Selection.ClearContents

Dim uln As Long

Sheets("PROJETO").Select

uln = Worksheets("PROJETO").Cells(Rows.Count, 5).End(xlUp).Row

Application.Calculation = xlManual

    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=IF(XLOOKUP(RC6,TAB_PRODUTO[CÓDIGO],TAB_PRODUTO[DESCRICAO_COMPLETA],""NÃO INFORMADO"",0)=""                              "",""NÃO INFORMADO"",IF(XLOOKUP(RC6,TAB_PRODUTO[CÓDIGO],TAB_PRODUTO[DESCRICAO_COMPLETA],""NÃO INFORMADO"",0)=0,""NÃO INFORMADO"",XLOOKUP(RC6,TAB_PRODUTO[CÓDIGO],TAB_PRODUTO[DESCRICAO_COMPLETA],""NÃO INFORMADO"",0)))"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=IF(XLOOKUP(RC6,TAB_PRODUTO[CÓDIGO],TAB_PRODUTO[TIPO],""NÃO INFORMADO"",0)=""                              "",""NÃO INFORMADO"",IF(XLOOKUP(RC6,TAB_PRODUTO[CÓDIGO],TAB_PRODUTO[TIPO],""NÃO INFORMADO"",0)=0,""NÃO INFORMADO"",XLOOKUP(RC6,TAB_PRODUTO[CÓDIGO],TAB_PRODUTO[TIPO],""NÃO INFORMADO"",0)))"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=IF(XLOOKUP(RC6,TAB_PRODUTO[CÓDIGO],TAB_PRODUTO[MEDIDA],""NÃO INFORMADO"",0)=""                              "",""NÃO INFORMADO"",IF(XLOOKUP(RC6,TAB_PRODUTO[CÓDIGO],TAB_PRODUTO[MEDIDA],""NÃO INFORMADO"",0)=0,""NÃO INFORMADO"",XLOOKUP(RC6,TAB_PRODUTO[CÓDIGO],TAB_PRODUTO[MEDIDA],""NÃO INFORMADO"",0)))"
    
'Range("H2:I2").Select
'
'Selection.AutoFill Destination:=Range(Cells(2, 8), Cells(uln, 9))

Application.Calculation = xlAutomatic
    
Range(Cells(2, 8), Cells(uln, 9)).Select

Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
Application.Calculation = xlManual
    
'Range("K2").Select
'
'Selection.AutoFill Destination:=Range(Cells(2, 11), Cells(uln, 11))
'
Application.Calculation = xlAutomatic

Range(Cells(2, 11), Cells(uln, 11)).Select

Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range(loc).Select
    
    Range("TAB_BASELINE[ESPECIFICAÇÕES]").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="PRONTO,AJUSTE,DESCONSIDERAR"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "ESPECIFICAÇÕES:"
        .ErrorTitle = ""
        .InputMessage = _
        "" & Chr(10) & "PRONTO = Nada muda na geometria da peça" & Chr(10) & "AJUSTE = Geometria da peça passivo de alteração. PN será alterado" & Chr(10) & "DESCONSIDERAR = Part Number não aplicável para estre projeto"
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("TAB_BASELINE[TOTVs]").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="OK,ÑOK"
        .IgnoreBlank = False
        .InCellDropdown = True
        .InputTitle = "TOTVs:"
        .ErrorTitle = "ANTEÇÃO"
        .InputMessage = _
        "" & Chr(10) & "OK = Código e Estrutura minima cadastrada no TOTVs" & Chr(10) & "" & Chr(10) & "Vazio ou ÑOK = Cadastro não concluído"
        .ErrorMessage = _
        "Por favor Selecionar uma das opções da lista suspensa na célula."
        .ShowInput = True
        .ShowError = True
    End With
    Range("TAB_BASELINE[QTDE]").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="OK,NOK"
        .IgnoreBlank = False
        .InCellDropdown = True
        .InputTitle = "QTDE:"
        .ErrorTitle = "ANTEÇÃO"
        .InputMessage = _
        "" & Chr(10) & "OK = Código e Estrutura minima cadastrada no TOTVs" & Chr(10) & "" & Chr(10) & "Vazio ou ÑOK = Cadastro não concluído"
        .ErrorMessage = _
        "Por favor Selecionar uma das opções da lista suspensa na célula."
        .ShowInput = True
        .ShowError = True
    End With
    
Application.ScreenUpdating = True

uln = Empty

Sheets("PROJETO").Range("I1").Select

End If

End Sub

Sub UPDATE_PRODUTO()
    Dim tb As ListObject
    Set tb = Planilha25.ListObjects("TAB_PRODUTO")
    tb.Refresh
End Sub

Sub validação()
    Dim tb As ListObject
    Set tb = Planilha1.ListObjects("TAB_BASELINE_VALIDAÇÃO")
    tb.Refresh
End Sub

Sub Endereco()
    Dim agora As String
    agora = ActiveCell.Address
End Sub


Sub COMENTARIOS()

Application.ScreenUpdating = False

Dim myCmt As CommentThreaded
Dim myRp As CommentThreaded
Dim curwks As Worksheet
Dim newwks As Worksheet
Dim myList As ListObject
Dim i As Long
Dim iR As Long
Dim iRCol As Long
Dim ListCols As Long
Dim cmtCount As Long
Dim WS_Count As Integer
Dim p As Integer

WS_Count = ActiveWorkbook.Worksheets.Count

Set newwks = Worksheets("COMENTARIOS")

 newwks.Range("A1:H1").Value = Array("PLANILHA", "ENDEREÇO", "LINK", "AUTOR", "DATA", "RESPOSTAS", "RESOLVIDO", "COMENTARIO")
 
 i = 1

For p = 1 To WS_Count

Set curwks = Worksheets(p)
cmtCount = curwks.CommentsThreaded.Count

If cmtCount > 0 Then

For Each myCmt In curwks.CommentsThreaded
   With newwks
     i = i + 1
     On Error Resume Next
     .Cells(i, 1).Value = curwks.Name
     .Cells(i, 2).Value = myCmt.Parent.Address
'     .Cells(I, 3).Value =
     .Cells(i, 4).Value = myCmt.Author.Name
     .Cells(i, 5).Value = myCmt.Date
     .Cells(i, 6).Value = myCmt.Replies.Count
     .Cells(i, 7).Value = myCmt.Resolved
     .Cells(i, 8).Value = myCmt.Text
   If myCmt.Replies.Count >= 1 Then
    iR = 1
    iRCol = 9
    For iR = 1 To myCmt.Replies.Count
      .Cells(1, iRCol).Value = "RESPOSTAS " & iR
      .Cells(i, iRCol).Value _
        = myCmt.Replies(iR).Author.Name _
          & vbCrLf _
          & myCmt.Replies(iR).Date _
          & vbCrLf _
          & myCmt.Replies(iR).Text
      iRCol = iRCol + 1
    Next iR
   End If
   End With
Next myCmt

End If

Next p

On Error Resume Next
With newwks
.ListObjects.Add(xlSrcRange, _
  .Cells(1, 1) _
    .CurrentRegion, , xlYes) _
    .Name = ""
End With

Set myList = newwks.ListObjects(1)
myList.TableStyle = "TableStyleDark10"
ListCols = myList.DataBodyRange _
  .Columns.Count

With myList.DataBodyRange
  .Cells.VerticalAlignment = xlTop
  .Columns.EntireColumn.ColumnWidth = 50
  .Cells.WrapText = True
  .Columns.EntireColumn.AutoFit
  .Rows.EntireRow.AutoFit
End With

Rows(1).Font.Color = RGB(0, 0, 0)

Range("C2").Formula2R1C1 = "=HYPERLINK(CONCAT(""#"",[@Planilha],""!"",[@Endereço]),""LINK"")"

Application.ScreenUpdating = True

myCmt = Nothing
myRp = Nothing
curwks = Nothing
newwks Nothing
'myList = Nothing
i = Empty
iR = Empty
iRCol = Empty
ListCols = Empty
cmtCount = Empty
WS_Count = Empty
p = Empty

End Sub


Public Sub UpdatePowerQueries()

Range("P1").Value = Now

Dim lTest As Long, cn As WorkbookConnection
On Error Resume Next
For Each cn In ThisWorkbook.Connections
lTest = InStr(1, cn.OLEDBConnection.Connection, "Provider=Microsoft.Mashup.OleDb.1", vbTextCompare)
If Err.Number <> 0 Then
Err.Clear
Exit For
End If
If lTest > 0 Then cn.Refresh
Debug.Print cn
Next cn

Range("P2").Value = Now

End Sub

Sub planilhas()

Dim aba As Worksheet
Dim x As Integer
x = 1

For Each aba In ThisWorkbook.Sheets

Sheets("PAINEL").Cells(x, 1).Value = aba.Name
x = x + 1

Next

End Sub

Sub TAB_PROJETO()

Application.ScreenUpdating = False

Dim uln As Long
Dim xList As ListObject

Sheets("PROJETO").Select

uln = Worksheets("PROJETO").Cells(Rows.Count, 5).End(xlUp).Row

    Range(Cells(2, 1), Cells(uln, 32)).Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

For Each xList In Sheets("PROJETO").ListObjects
    xList.Unlist
Next

Sheets("PROJETO").ListObjects.Add(xlSrcRange, Range("$A$1:$AF$" & uln), , xlYes).Name = "TAB_BASELINE"

ActiveSheet.ListObjects("TAB_BASELINE").TableStyle = "TableStyleMedium4"

Application.ScreenUpdating = True

End Sub

Sub ORCAMENTO()
Dim uln As Long
Dim x As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    Range(Cells(2, 12), Cells(1000000, 38)).ClearContents
    Range("AY2").ClearContents
    Range(Cells(2, 55), Cells(1000000, 55)).ClearContents
    Rows("2:50000").Interior.Pattern = xlNone
    Rows("1:1").RowHeight = 100


uln = Planilha4.Cells(Rows.Count, 6).End(xlUp).Row

    Planilha4.Select
    Range("L1").Select
    Range("L1").FormulaR1C1 = "PER/UNIT"
    Range("M1").FormulaR1C1 = "ULTIMA ATUALIZAÇÃO EM DIAS"
    Range("N1").FormulaR1C1 = "VALOR TOTAL"
    Range("O1").FormulaR1C1 = "TAXA"
    Range("P1").FormulaR1C1 = "VALOR TOTAL C/ TAXA AJUSTE"
    Range("Q1").FormulaR1C1 = "VALOR POR PN"
    Range("R1").FormulaR1C1 = "0"
    Range("S1").FormulaR1C1 = "1"
    Range("T1").FormulaR1C1 = "2"
    Range("U1").FormulaR1C1 = "3"
    Range("V1").FormulaR1C1 = "4"
    Range("W1").FormulaR1C1 = "5"
    Range("X1").FormulaR1C1 = "6"
    Range("Y1").FormulaR1C1 = "7"
    Range("Z1").FormulaR1C1 = "8"
    Range("AA1").FormulaR1C1 = "9"
    Range("AB1").FormulaR1C1 = "0"
    Range("AC1").FormulaR1C1 = "1"
    Range("AD1").FormulaR1C1 = "2"
    Range("AE1").FormulaR1C1 = "3"
    Range("AF1").FormulaR1C1 = "4"
    Range("AG1").FormulaR1C1 = "5"
    Range("AH1").FormulaR1C1 = "6"
    Range("AI1").FormulaR1C1 = "7"
    Range("AJ1").FormulaR1C1 = "8"
    Range("AK1").FormulaR1C1 = "9"
    Range("AL1").FormulaR1C1 = "=CONCAT(COUNTIFS(C[-32],""MP"",C[-29],"""")&"" - MP SEM HISTÓRICO DO VALOR"",CHAR(10),TEXT(R[1]C,""R$ #.##0,00""),"" - VALOR DA ESTRUTURA"",CHAR(10),R[2]C,"" - DIAS ENTREGA MP"",CHAR(10),TEXT(TEMPO_ESTIMADO!R[11]C[-31],""0""),"" - DIAS ÚTEIS DESENVOLVIMENTO"")"
    Range("AL2").FormulaR1C1 = "=SUM(C[-22])"
    Range("AL3").FormulaR1C1 = "=MAX(C[-28])"
    Range("L2").FormulaR1C1 = "=IF(RC[-6]=""MP"",IF(RC[-3]="""",XLOOKUP(RC[-9],C[39],C[43],""SEM INFORMAÇÃO"",0),RC[-3]),"""")"
    Range("M2").FormulaR1C1 = "=IF(RC[-7]=""MP"",IF(RC[-6]="""",""SEM INFORMAÇÃO"",TODAY()-RC[-6]),"""")"
    Range("N2").FormulaR1C1 = "=IF(OR(RC[-2]="""",RC[-2]=""SEM INFORMAÇÃO""),"""",RC[-6]*RC[-2])"
'    Range("O2").FormulaR1C1 = "100%"
    Range("P2").FormulaR1C1 = "=IFERROR(RC[-2]*RC[-1],"""")"
    Range("Q2").FormulaR1C1 = "=SUM(RC[1]:RC[10])"
    Range("R2").FormulaR1C1 = "=IF(RC[10]>0,IF(RC[10]<>R[-1]C[10],SUMIFS(C16,C[10],RC[10]),""""),0)"
    Range("S2").FormulaR1C1 = "=IF(RC[10]>0,IF(RC[10]<>R[-1]C[10],SUMIFS(C16,C[10],RC[10]),""""),0)"
    Range("T2").FormulaR1C1 = "=IF(RC[10]>0,IF(RC[10]<>R[-1]C[10],SUMIFS(C16,C[10],RC[10]),""""),0)"
    Range("U2").FormulaR1C1 = "=IF(RC[10]>0,IF(RC[10]<>R[-1]C[10],SUMIFS(C16,C[10],RC[10]),""""),0)"
    Range("V2").FormulaR1C1 = "=IF(RC[10]>0,IF(RC[10]<>R[-1]C[10],SUMIFS(C16,C[10],RC[10]),""""),0)"
    Range("W2").FormulaR1C1 = "=IF(RC[10]>0,IF(RC[10]<>R[-1]C[10],SUMIFS(C16,C[10],RC[10]),""""),0)"
    Range("X2").FormulaR1C1 = "=IF(RC[10]>0,IF(RC[10]<>R[-1]C[10],SUMIFS(C16,C[10],RC[10]),""""),0)"
    Range("Y2").FormulaR1C1 = "=IF(RC[10]>0,IF(RC[10]<>R[-1]C[10],SUMIFS(C16,C[10],RC[10]),""""),0)"
    Range("Z2").FormulaR1C1 = "=IF(RC[10]>0,IF(RC[10]<>R[-1]C[10],SUMIFS(C16,C[10],RC[10]),""""),0)"
    Range("AA2").FormulaR1C1 = "=IF(RC[10]>0,IF(RC[10]<>R[-1]C[10],SUMIFS(C16,C[10],RC[10]),""""),0)"
    Range("AB2").FormulaR1C1 = "=COUNTIF(R2C2:RC2,R1C)"
    Range("AC2").FormulaR1C1 = "=COUNTIF(R2C2:RC2,R1C)"
    Range("AD2").FormulaR1C1 = "=COUNTIF(R2C2:RC2,R1C)"
    Range("AE2").FormulaR1C1 = "=COUNTIF(R2C2:RC2,R1C)"
    Range("AF2").FormulaR1C1 = "=COUNTIF(R2C2:RC2,R1C)"
    Range("AG2").FormulaR1C1 = "=COUNTIF(R2C2:RC2,R1C)"
    Range("AH2").FormulaR1C1 = "=COUNTIF(R2C2:RC2,R1C)"
    Range("AI2").FormulaR1C1 = "=COUNTIF(R2C2:RC2,R1C)"
    Range("AJ2").FormulaR1C1 = "=COUNTIF(R2C2:RC2,R1C)"
    Range("AK2").FormulaR1C1 = "=COUNTIF(R2C2:RC2,R1C)"
    Range("AY2").Formula2R1C1 = "=UNIQUE(FILTER(C[-48]:C[-45],(C[-40]="""")*(C[-45]=""MP"")))"
    
    Range("AN4").Select
    
    Range("N2").Select
    
    If uln > 2 Then
    
    Range("L2:AK2").Select

    Selection.AutoFill Destination:=Range(Cells(2, 12), Cells(uln, 37)), Type:=xlFillDefault
    
'    Range("O2").FormulaR1C1 = "100%"
'    Range("O3").FormulaR1C1 = "100%"
    Range("P2").FormulaR1C1 = "=IFERROR(RC[-2]*RC[-1],"""")"
    Range("P3").FormulaR1C1 = "=IFERROR(RC[-2]*RC[-1],"""")"
    
    Range("P2:P3").Select

    Selection.AutoFill Destination:=Range(Cells(2, 16), Cells(uln, 16))
    
    Calculate
    
    Range(Cells(2, 12), Cells(uln, 37)).Select
    
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'    Application.CutCopyMode = False
    
    Calculate
    
    End If
    
    Columns("R:AL").Select
    Selection.EntireColumn.Hidden = True
    
    For x = 2 To uln
        If Cells(x, 13).Value = "SEM INFORMAÇÃO" Then
            Range(Cells(x, 1), Cells(x, 14)).Interior.Color = 65535
        End If
    Next x
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    
    Range("O2").Select
    
End Sub

Sub FORMT_CONDICIONAL()
Dim MyRange As Range
Dim uln As Long
uln = Worksheets("PROJETO").Cells(Rows.Count, 5).End(xlUp).Row

Set MyRange = Range("A2:AF" & uln)

MyRange.FormatConditions.Delete

MyRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=$O2=""TEMPORÁRIO"""
MyRange.FormatConditions(1).Interior.Color = RGB(0, 0, 0)
MyRange.FormatConditions(1).Font.Color = RGB(255, 255, 255)

MyRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=$O2=""DESCONSIDERAR"""
MyRange.FormatConditions(2).Interior.Color = RGB(191, 191, 191)
MyRange.FormatConditions(2).Font.Color = RGB(89, 89, 89)

MyRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=$O2=""SUBSTITUÍDO"""
MyRange.FormatConditions(3).Interior.Color = RGB(191, 191, 191)
MyRange.FormatConditions(3).Font.Color = RGB(89, 89, 89)

MyRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=$AD2=0"
MyRange.FormatConditions(4).Interior.Color = RGB(191, 191, 191)
MyRange.FormatConditions(4).Font.Color = RGB(89, 89, 89)

MyRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=$AE2=""FINALIZADO"""
MyRange.FormatConditions(5).Interior.Color = RGB(198, 239, 206)
MyRange.FormatConditions(5).Font.Color = RGB(0, 97, 0)

MyRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=$AE2=""PENDENTE"""
MyRange.FormatConditions(6).Interior.Color = RGB(255, 199, 206)
MyRange.FormatConditions(6).Font.Color = RGB(156, 0, 6)

MyRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=$O2=""AJUSTE"""
MyRange.FormatConditions(7).Interior.Color = RGB(255, 235, 156)
MyRange.FormatConditions(7).Font.Color = RGB(156, 87, 0)

MyRange.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=$O2=""PRONTO"""
MyRange.FormatConditions(8).Interior.Color = RGB(0, 255, 0)
MyRange.FormatConditions(8).Font.Color = RGB(55, 86, 35)

End Sub




