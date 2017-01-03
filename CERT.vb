Option compatible

sub cert4
dim document   as object
dim dispatcher as object
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ==============================================================

Dim oSheets
dim plan1(0) as new com.sun.star.beans.PropertyValue
plan1(0).Name = "Nr"
plan1(0).Value = 1
dim args1(0) as new com.sun.star.beans.PropertyValue
args1(0).Name = "Sel"
args1(0).Value = false
dim args7(1) as new com.sun.star.beans.PropertyValue
args7(0).Name = "By"
args7(0).Value = 1
args7(1).Name = "Sel"
args7(1).Value = false
dim args2(0) as new com.sun.star.beans.PropertyValue
args2(0).Name = "By"
args2(0).Value = 4
dim args3(1) as new com.sun.star.beans.PropertyValue
args3(0).Name = "By"
args3(0).Value = 3
args3(1).Name = "Sel"
args3(1).Value = false

rem--------------------------Aruma datas e horarios--------------------

for n = 0 to 500
dispatcher.executeDispatch(document, ".uno:GoToStart", "", 0, args1())
oCell = ThisComponent.Sheets(0).getCellByPosition(4, n)
a = oCell.getString()
oCell = ThisComponent.Sheets(0).getCellByPosition(3, n)
b = oCell.getString()
oCell = ThisComponent.Sheets(0).getCellByPosition(0, n)
c = oCell.getString()
if a ="" and c<>"" then
args3(0).Value = (n)
dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args3())
args7(0).Value = 1
dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args7())
dispatcher.executeDispatch(document, ".uno:GoRightSel", "", 0, args2())
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())
args3(0).Value = 1
if b = "" then
args3(0).Value = 2
endif
dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args3())
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())
endif
next n

rem-------------------------insere nova tabela-----------------------------

oSheets = ThisComponent.Sheets
oSheets.insertNewByName("Mes", 1)
dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, plan1())
dispatcher.executeDispatch(document, ".uno:GoToStart", "", 0, args1())

rem------------------------retira todos os testes----------------------------

dim args9(17) as new com.sun.star.beans.PropertyValue
args9(0).Name = "SearchItem.StyleFamily"
args9(0).Value = 2
args9(1).Name = "SearchItem.CellType"
args9(1).Value = 0
args9(2).Name = "SearchItem.RowDirection"
args9(2).Value = true
args9(3).Name = "SearchItem.AllTables"
args9(3).Value = false
args9(4).Name = "SearchItem.Backward"
args9(4).Value = false
args9(5).Name = "SearchItem.Pattern"
args9(5).Value = false
args9(6).Name = "SearchItem.Content"
args9(6).Value = false
args9(7).Name = "SearchItem.AsianOptions"
args9(7).Value = false
args9(8).Name = "SearchItem.AlgorithmType"
args9(8).Value = 0
args9(9).Name = "SearchItem.SearchFlags"
args9(9).Value = 0
args9(10).Name = "SearchItem.SearchString"
args9(10).Value = "SP_RFB_SRRF_TESTE1_SAO_PAULO_C22801"
args9(11).Name = "SearchItem.ReplaceString"
args9(11).Value = "M"
args9(12).Name = "SearchItem.Locale"
args9(12).Value = 255
args9(13).Name = "SearchItem.ChangedChars"
args9(13).Value = 2
args9(14).Name = "SearchItem.DeletedChars"
args9(14).Value = 2
args9(15).Name = "SearchItem.InsertedChars"
args9(15).Value = 2
args9(16).Name = "SearchItem.TransliterateFlags"
args9(16).Value = 256
args9(17).Name = "SearchItem.Command"
args9(17).Value = 3
dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args9())
dispatcher.executeDispatch(document, ".uno:GoToStart", "", 0, args1())

rem------------------------------------------------------------------------

dim args6(17) as new com.sun.star.beans.PropertyValue
args6(0).Name = "SearchItem.StyleFamily"
args6(0).Value = 2
args6(1).Name = "SearchItem.CellType"
args6(1).Value = 0
args6(2).Name = "SearchItem.RowDirection"
args6(2).Value = true
args6(3).Name = "SearchItem.AllTables"
args6(3).Value = false
args6(4).Name = "SearchItem.Backward"
args6(4).Value = false
args6(5).Name = "SearchItem.Pattern"
args6(5).Value = false
args6(6).Name = "SearchItem.Content"
args6(6).Value = false
args6(7).Name = "SearchItem.AsianOptions"
args6(7).Value = false
args6(8).Name = "SearchItem.AlgorithmType"
args6(8).Value = 0
args6(9).Name = "SearchItem.SearchFlags"
args6(9).Value = 0
args6(10).Name = "SearchItem.SearchString"
args6(10).Value = "SP"
args6(11).Name = "SearchItem.ReplaceString"
args6(11).Value = ""
args6(12).Name = "SearchItem.Locale"
args6(12).Value = 255
args6(13).Name = "SearchItem.ChangedChars"
args6(13).Value = 2
args6(14).Name = "SearchItem.DeletedChars"
args6(14).Value = 2
args6(15).Name = "SearchItem.InsertedChars"
args6(15).Value = 2
args6(16).Name = "SearchItem.TransliterateFlags"
args6(16).Value = 0
args6(17).Name = "SearchItem.Command"
args6(17).Value = 0
for m = 0 to 500
dispatcher.executeDispatch(document, ".uno:GoToStart", "", 0, args1())
dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args6())
oCell = ThisComponent.getCurrentSelection()
d = oCell.getString()
if d<>"iniciando" then
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
plan1(0).Value = 2
dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, plan1())
dispatcher.executeDispatch(document, ".uno:GoToStart", "", 0, args1())
args7(0).Value = (m)
dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args7())
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())
plan1(0).Value = 1
dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, plan1())
dispatcher.executeDispatch(document, ".uno:GoToStart", "", 0, args1())
dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args6())
args7(0).Value = 1
dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args7())
args7(0).Value = 3
dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args7())
args2(0).Value = 2
dispatcher.executeDispatch(document, ".uno:GoRightSel", "", 0, args2())
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
plan1(0).Value = 2
dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, plan1())
dispatcher.executeDispatch(document, ".uno:GoToStart", "", 0, args1())
args7(0).Value = (m)
dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args7())
args7(0).Value = 1
dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args7())
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())
plan1(0).Value = 1
dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, plan1())
dispatcher.executeDispatch(document, ".uno:GoToStart", "", 0, args1())
dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args6())
args7(0).Value = 3
dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args7())
args2(0).Value = 2
dispatcher.executeDispatch(document, ".uno:GoRightSel", "", 0, args2())
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())
plan1(0).Value = 2
dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, plan1())
dispatcher.executeDispatch(document, ".uno:GoToStart", "", 0, args1())
args7(0).Value = (m)
dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args7())
args7(0).Value = 3
dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args7())
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())
plan1(0).Value = 1
dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, plan1())
dispatcher.executeDispatch(document, ".uno:GoToStart", "", 0, args1())
dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args6())
dispatcher.executeDispatch(document, ".uno:ClearContents", "", 0, Array())
endif
next m
plan1(0).Value = 2
dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, plan1())

dim vari1(0) as new com.sun.star.beans.PropertyValue
vari1(0).Name = "Sel"
vari1(0).Value = false

dispatcher.executeDispatch(document, ".uno:GoToStart", "", 0, vari1())

rem ----------------------------------------------------------------------
dim vari2(1) as new com.sun.star.beans.PropertyValue
vari2(0).Name = "By"
vari2(0).Value = 1
vari2(1).Name = "Sel"
vari2(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari2())

rem ----------------------------------------------------------------------
dim vari3(0) as new com.sun.star.beans.PropertyValue
vari3(0).Name = "Sel"
vari3(0).Value = true

dispatcher.executeDispatch(document, ".uno:GoToEndOfData", "", 0, vari3())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim vari5(1) as new com.sun.star.beans.PropertyValue
vari5(0).Name = "By"
vari5(0).Value = 1
vari5(1).Name = "Sel"
vari5(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari5())

rem ----------------------------------------------------------------------
dim vari6(1) as new com.sun.star.beans.PropertyValue
vari6(0).Name = "By"
vari6(0).Value = 1
vari6(1).Name = "Sel"
vari6(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari6())

rem ----------------------------------------------------------------------
dim vari7(1) as new com.sun.star.beans.PropertyValue
vari7(0).Name = "By"
vari7(0).Value = 1
vari7(1).Name = "Sel"
vari7(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari7())

rem ----------------------------------------------------------------------
dim vari8(1) as new com.sun.star.beans.PropertyValue
vari8(0).Name = "By"
vari8(0).Value = 1
vari8(1).Name = "Sel"
vari8(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari8())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())

rem ----------------------------------------------------------------------
dim vari10(1) as new com.sun.star.beans.PropertyValue
vari10(0).Name = "By"
vari10(0).Value = 1
vari10(1).Name = "Sel"
vari10(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari10())

rem ----------------------------------------------------------------------
dim vari11(1) as new com.sun.star.beans.PropertyValue
vari11(0).Name = "By"
vari11(0).Value = 1
vari11(1).Name = "Sel"
vari11(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari11())

rem ----------------------------------------------------------------------
dim vari12(0) as new com.sun.star.beans.PropertyValue
vari12(0).Name = "By"
vari12(0).Value = 1

dispatcher.executeDispatch(document, ".uno:GoDownToEndOfDataSel", "", 0, vari12())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim vari14(1) as new com.sun.star.beans.PropertyValue
vari14(0).Name = "By"
vari14(0).Value = 1
vari14(1).Name = "Sel"
vari14(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari14())

rem ----------------------------------------------------------------------
dim vari15(1) as new com.sun.star.beans.PropertyValue
vari15(0).Name = "By"
vari15(0).Value = 1
vari15(1).Name = "Sel"
vari15(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari15())

rem ----------------------------------------------------------------------
dim vari16(1) as new com.sun.star.beans.PropertyValue
vari16(0).Name = "By"
vari16(0).Value = 1
vari16(1).Name = "Sel"
vari16(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari16())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())

rem ----------------------------------------------------------------------
dim vari18(1) as new com.sun.star.beans.PropertyValue
vari18(0).Name = "By"
vari18(0).Value = 1
vari18(1).Name = "Sel"
vari18(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari18())

rem ----------------------------------------------------------------------
dim vari19(1) as new com.sun.star.beans.PropertyValue
vari19(0).Name = "By"
vari19(0).Value = 1
vari19(1).Name = "Sel"
vari19(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari19())

rem ----------------------------------------------------------------------
dim vari20(1) as new com.sun.star.beans.PropertyValue
vari20(0).Name = "By"
vari20(0).Value = 1
vari20(1).Name = "Sel"
vari20(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari20())

rem ----------------------------------------------------------------------
dim vari21(1) as new com.sun.star.beans.PropertyValue
vari21(0).Name = "By"
vari21(0).Value = 1
vari21(1).Name = "Sel"
vari21(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari21())

rem ----------------------------------------------------------------------
dim vari22(0) as new com.sun.star.beans.PropertyValue
vari22(0).Name = "By"
vari22(0).Value = 1

dispatcher.executeDispatch(document, ".uno:GoDownSel", "", 0, vari22())

rem ----------------------------------------------------------------------
dim vari23(1) as new com.sun.star.beans.PropertyValue
vari23(0).Name = "By"
vari23(0).Value = 1
vari23(1).Name = "Sel"
vari23(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoUp", "", 0, vari23())

rem ----------------------------------------------------------------------
dim vari24(0) as new com.sun.star.beans.PropertyValue
vari24(0).Name = "By"
vari24(0).Value = 1

dispatcher.executeDispatch(document, ".uno:GoDownToEndOfDataSel", "", 0, vari24())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim vari26(1) as new com.sun.star.beans.PropertyValue
vari26(0).Name = "By"
vari26(0).Value = 1
vari26(1).Name = "Sel"
vari26(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari26())

rem ----------------------------------------------------------------------
dim vari27(1) as new com.sun.star.beans.PropertyValue
vari27(0).Name = "By"
vari27(0).Value = 1
vari27(1).Name = "Sel"
vari27(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari27())

rem ----------------------------------------------------------------------
dim vari28(1) as new com.sun.star.beans.PropertyValue
vari28(0).Name = "By"
vari28(0).Value = 1
vari28(1).Name = "Sel"
vari28(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari28())

rem ----------------------------------------------------------------------
dim vari29(1) as new com.sun.star.beans.PropertyValue
vari29(0).Name = "By"
vari29(0).Value = 1
vari29(1).Name = "Sel"
vari29(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari29())

rem ----------------------------------------------------------------------
dim vari30(1) as new com.sun.star.beans.PropertyValue
vari30(0).Name = "By"
vari30(0).Value = 1
vari30(1).Name = "Sel"
vari30(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari30())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())

rem ----------------------------------------------------------------------
dim vari32(1) as new com.sun.star.beans.PropertyValue
vari32(0).Name = "By"
vari32(0).Value = 1
vari32(1).Name = "Sel"
vari32(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari32())

rem ----------------------------------------------------------------------
dim vari33(1) as new com.sun.star.beans.PropertyValue
vari33(0).Name = "By"
vari33(0).Value = 1
vari33(1).Name = "Sel"
vari33(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari33())

rem ----------------------------------------------------------------------
dim vari34(1) as new com.sun.star.beans.PropertyValue
vari34(0).Name = "By"
vari34(0).Value = 1
vari34(1).Name = "Sel"
vari34(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari34())

rem ----------------------------------------------------------------------
dim vari35(1) as new com.sun.star.beans.PropertyValue
vari35(0).Name = "By"
vari35(0).Value = 1
vari35(1).Name = "Sel"
vari35(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari35())

rem ----------------------------------------------------------------------
dim vari36(1) as new com.sun.star.beans.PropertyValue
vari36(0).Name = "By"
vari36(0).Value = 1
vari36(1).Name = "Sel"
vari36(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari36())

rem ----------------------------------------------------------------------
dim vari37(1) as new com.sun.star.beans.PropertyValue
vari37(0).Name = "By"
vari37(0).Value = 1
vari37(1).Name = "Sel"
vari37(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari37())

rem ----------------------------------------------------------------------
dim vari38(0) as new com.sun.star.beans.PropertyValue
vari38(0).Name = "By"
vari38(0).Value = 1

dispatcher.executeDispatch(document, ".uno:GoDownToEndOfDataSel", "", 0, vari38())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim vari40(1) as new com.sun.star.beans.PropertyValue
vari40(0).Name = "By"
vari40(0).Value = 1
vari40(1).Name = "Sel"
vari40(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari40())

rem ----------------------------------------------------------------------
dim vari41(1) as new com.sun.star.beans.PropertyValue
vari41(0).Name = "By"
vari41(0).Value = 1
vari41(1).Name = "Sel"
vari41(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari41())

rem ----------------------------------------------------------------------
dim vari42(1) as new com.sun.star.beans.PropertyValue
vari42(0).Name = "By"
vari42(0).Value = 1
vari42(1).Name = "Sel"
vari42(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari42())

rem ----------------------------------------------------------------------
dim vari43(1) as new com.sun.star.beans.PropertyValue
vari43(0).Name = "By"
vari43(0).Value = 1
vari43(1).Name = "Sel"
vari43(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari43())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())

rem ----------------------------------------------------------------------
dim vari45(1) as new com.sun.star.beans.PropertyValue
vari45(0).Name = "By"
vari45(0).Value = 1
vari45(1).Name = "Sel"
vari45(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari45())

rem ----------------------------------------------------------------------
dim vari46(1) as new com.sun.star.beans.PropertyValue
vari46(0).Name = "By"
vari46(0).Value = 1
vari46(1).Name = "Sel"
vari46(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari46())

rem ----------------------------------------------------------------------
dim vari47(1) as new com.sun.star.beans.PropertyValue
vari47(0).Name = "By"
vari47(0).Value = 1
vari47(1).Name = "Sel"
vari47(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari47())

rem ----------------------------------------------------------------------
dim vari48(1) as new com.sun.star.beans.PropertyValue
vari48(0).Name = "By"
vari48(0).Value = 1
vari48(1).Name = "Sel"
vari48(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari48())

rem ----------------------------------------------------------------------
dim vari49(1) as new com.sun.star.beans.PropertyValue
vari49(0).Name = "By"
vari49(0).Value = 1
vari49(1).Name = "Sel"
vari49(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari49())

rem ----------------------------------------------------------------------
dim vari50(1) as new com.sun.star.beans.PropertyValue
vari50(0).Name = "By"
vari50(0).Value = 1
vari50(1).Name = "Sel"
vari50(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, vari50())

rem ----------------------------------------------------------------------
dim vari51(1) as new com.sun.star.beans.PropertyValue
vari51(0).Name = "By"
vari51(0).Value = 1
vari51(1).Name = "Sel"
vari51(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari51())

rem ----------------------------------------------------------------------
dim vari52(0) as new com.sun.star.beans.PropertyValue
vari52(0).Name = "By"
vari52(0).Value = 1

dispatcher.executeDispatch(document, ".uno:GoDownToEndOfDataSel", "", 0, vari52())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim vari54(1) as new com.sun.star.beans.PropertyValue
vari54(0).Name = "By"
vari54(0).Value = 1
vari54(1).Name = "Sel"
vari54(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari54())

rem ----------------------------------------------------------------------
dim vari55(1) as new com.sun.star.beans.PropertyValue
vari55(0).Name = "By"
vari55(0).Value = 1
vari55(1).Name = "Sel"
vari55(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari55())

rem ----------------------------------------------------------------------
dim vari56(1) as new com.sun.star.beans.PropertyValue
vari56(0).Name = "By"
vari56(0).Value = 1
vari56(1).Name = "Sel"
vari56(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari56())

rem ----------------------------------------------------------------------
dim vari57(1) as new com.sun.star.beans.PropertyValue
vari57(0).Name = "By"
vari57(0).Value = 1
vari57(1).Name = "Sel"
vari57(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, vari57())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())
end sub