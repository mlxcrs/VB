Option compatible

sub Principal
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
dim args1(1) as new com.sun.star.beans.PropertyValue
dim args2(0) as new com.sun.star.beans.PropertyValue
dim args3(0) as new com.sun.star.beans.PropertyValue
dim args4(0) as new com.sun.star.beans.PropertyValue
dim args5(1) as new com.sun.star.beans.PropertyValue
dim args6(0) as new com.sun.star.beans.PropertyValue
dim args7(0) as new com.sun.star.beans.PropertyValue
Dim oDescriptor 'The search descriptor
Dim oResultado 'The found range
Dim i%
Dim j%
Dim b%
Dim planilha0%, planilha%, planilha2%
Dim Decisao
Dim coluna (0) as new com.sun.star.beans.PropertyValue
Dim coluna1 as String
dim palavra as string
Dim tamanho as integer
Dim oProgressBar as Object, oProgressBarModel As Object, oDialog as Object
Dim ProgressValue As Long
DialogLibraries.loadLibrary("Standard")
oDialog = CreateUnoDialog(DialogLibraries.Standard.Dialog1)

i=0
j=1

oConv = Thiscomponent.createInstance("com.sun.star.table.CellAddressConversion")
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
rem ----------------------------------------------------------------------
rem declaração das variaveis e obtenção dos valores iniciais--------------
palavra = inputbox("Por qual palavra você deseja pesquisar?")
tamanho = (len(palavra)/4)-1 ' faz a conta depois de inserir a palavra 
Line40:

coluna1 = inputbox("Em qual campo quer aplicar a busca? Veja a explicação abaixo"&CHR$(10)&CHR$(10)& "0 - para Nome UASG;"&CHR$(10)&"1 - para Nome UORG; ou"&CHR$(10)& "2 - para descrição.")
'planilha = inputbox("Digite o número da planilha em que será feita a pesquisa")
'planilha0 = planilha-1
'planilha2 = planilha+1
'coluna(0).value = coluna1+"2:"+coluna1+"100000"

rem ----------------------------------------------------------------------
rem get access to the document
'oConv = Thiscomponent.Sheets(0).createInstance("com.sun.star.table.CellAddressConversion")
if coluna1 = 0 or coluna1=1 or coluna1=2 then
if coluna1 = 0 then
coluna1 = 2
else
if coluna1 = 1 then
coluna1 = 4
else
coluna1 = 7
endif
endif
oDescriptor = ThisComponent.Sheets(0).Columns(coluna1).createSearchDescriptor()
With oDescriptor
.SearchString = palavra
.SearchSimilarity = true
.SearchRegularExpression = true
.SearchSimilarityRemove = tamanho
.SearchSimilarityAdd = tamanho
.SearchSimilarityExchange = tamanho
End With 
' Find the first one
oResultado = ThisComponent.Sheets(0).Columns (coluna1).findAll(oDescriptor)

if not isNull (oResultado) then
b = oResultado.getCount()
Decisao = Msgbox ("Foram localizadas " & b & " ocorrências." &CHR$(10)& "Isto levará aproximadamente: " & b\300 & " minuto(s)" & CHR$(10) & CHR$(10) & "Deseja Continuar?", 4, "Tempo de espera"
if Decisao = 6 then
REM set minimum and maximum progress value
oProgressBarModel = oDialog.getModel().getByName( "ProgressBar1" )
oProgressBarModel.setPropertyValue( "ProgressValueMin", 0)
oProgressBarModel.setPropertyValue( "ProgressValueMax", b)
REM show progress bar
oDialog.setVisible( True )


	args3(0).Name = "Nr"
	args3(0).Value = 2
	dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, args3())
	dispatcher.executeDispatch(document, ".uno:SelectAll", "", 0, Array())
	dispatcher.executeDispatch(document, ".uno:ClearContents", "", 0, Array())
	args3(0).Name = "Nr"
	args3(0).Value = 1
	dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, args3())
	args3(0).Name = "ToPoint"
	args3(0).Value = "$A$4:$V$4"
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args3())
	rem ----------------------------------------------------------------------
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	rem ----------------------------------------------------------------------
	args5(0).Name = "Nr"
	args5(0).Value = 2
	dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, args5())
	rem ----------------------------------------------------------------------
	args3(0).Name = "ToPoint"
	args3(0).Value = "$A$1:$V$1"
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args3())
	dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())
	args4(0).Name = "Sel"
	args4(0).Value = false
	dispatcher.executeDispatch(document, ".uno:GoToStart", "", 0, args4())
	rem ----------------------------------------------------------------------

	For n = 0 to oResultado.getCount()-1
	oProgressBarModel.setPropertyValue( "ProgressValue", n )
rem	oDialog.getModel().getByName("TextField1").text = "Pesquisando " & n & " de " & b &" ocorrências."
	oFound = oResultado.getByIndex(n)
	i = oFound.getrows().getCount()-1
	ThisComponent.CurrentController.Select(oFound)
	rem-----------------------------------------------------------------------
	args1(0).Name = "By"
	args1(0).Value = 26
	args1(1).Name = "Sel"
	args1(1).Value = false'
	dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args1())

	args5(0).Name = "By"
	args5(0).Value = i
	args5(1).Name = "Sel"
	args5(1).Value = true
	dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args5())

	rem ----------------------------------------------------------------------
	args2(0).Name = "By"
	args2(0).Value = 26
	dispatcher.executeDispatch(document, ".uno:GoRightSel", "", 0, args2())
	rem ----------------------------------------------------------------------
	dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
	rem-----------------------------------------------------------------------
	args3(0).Name = "Nr"
	args3(0).Value = 2
	dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, args3())
	rem control + home--------------------------------------------------------
	args4(0).Name = "Sel"
	args4(0).Value = false
	dispatcher.executeDispatch(document, ".uno:GoToStart", "", 0, args4())
	rem ----------------------------------------------------------------------
	args5(0).Name = "By"
	args5(0).Value = 1
	args5(1).Name = "Sel"
	args5(1).Value = true
	dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args5())

	args5(0).Name = "By"
	args5(0).Value = j
	args5(1).Name = "Sel"
	args5(1).Value = false
	dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args5())
	rem ----------------------------------------------------------------------
	dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())
	rem ----------------------------------------------------------------------
	args6(0).Name = "Nr"
	args6(0).Value = 1
	dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, args6())
	rem control + home--------------------------------------------------------
	j=j + oFound.getrows().getCount()
	next n
	args5(0).Name = "Nr"
	args5(0).Value = 2
	dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, args5())
	rem ----------------------------------------------------------------------
	args4(0).Name = "Sel"
	args4(0).Value = false
	dispatcher.executeDispatch(document, ".uno:GoToStart", "", 0, args4())
else
Msgbox ("Pequisa Cancelada",0)
Endif
else
Print "não obtivemos resultados para a palavra procurada"
endif
else
Msgbox ("Escolha apenas uma das opções", 0, "Aviso")
Goto Line40
endif
rem ----------------------------------------------------------------------
end sub 