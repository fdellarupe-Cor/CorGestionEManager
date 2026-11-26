<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
function cnumer(cdato)
   ldato = len(cdato)

   for kk=1 to ldato
       if isnumeric(mid(cdato,kk,1)) then
           cnumer = cnumer & mid(cdato,kk,1)
       end if
   next


end function

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  nPerfil = request.querystring("nPerfil")

'  sHTML = ofv.FormBody2

'   sHTML = ofv.FormHeader(session("FormN") & " - Test")

  sHTML = ofv.MenuHeader("#ffffff", "")


  set cn = ofv.conectar(ofv.strconn2)





 

' nStilo = " Style={font-family:" & nblkff & ";font-size:" & nblkfs & ";color:#993333;background-color:#ffffff;font-weight:bold;} "

nfontx = "</font>"


'Bloque de Defaults
ncolcant = 1
nBorde = false
nblkfc = "#993333"
nblkbc = "#ffffff"
nblkff = "Tahoma"
nblkfs = "9pt"

nctitfc = "#000077"
nctitbc = "#ffffff"
nctitff = "Tahoma"
nctitfs = "8pt"

ncconfc = "#336699"
ncconbc = "#ffffff"
ncconff = "Tahoma"
ncconfs = "8pt"

redim aColumnas(0)


    strsql = "Select * From Tabperfil Where Tabperfil = '" & nPerfil & "'"
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then
       nTitulo = rs2(1)
       ncolcant = rs2(7)
      if cint(ncolcant) > 10 then ncolcant = 10
nBorde = rs2(2) 
nblkfc = rs2(3)
nblkbc = rs2(4)
nblkff = replace(rs2(5)," ","-")
'nblkfs = rs2(6)
nblkfs = cnumer(rs2(6)) & "pt"

nctitfc = rs2(8)
nctitbc = rs2(9)
nctitff = replace(rs2(10)," ","-")
'nctitfs = rs2(11)
nctitfs = cnumer(rs2(11)) & "pt"


ncconfc = rs2(12)
ncconbc = rs2(13)
ncconff = replace(rs2(14)," ","-")
'ncconfs = rs2(15)
ncconfs = cnumer(rs2(15)) & "pt"

      redim preserve aColumnas(cint(ncolcant) - 1)
      redim preserve aColumnasa(cint(ncolcant) - 1)
      for ix = 0 to UBound(aColumnas)
        aColumnas(ix) = rs2(ix + 16)
        aColumnasa(ix) = rs2(ix + 26)
      next
    end if

    call ofv.cerrarconsulta(rs2)
xhtml = ""
    if nBorde then
    xHTML = xHTML & ofv.GenTabla("","","","600 border","","0","0","0","0","#ffffff","","","no")
    else
    xHTML = xHTML & ofv.GenTabla("","","","600","","0","0","0","0","#ffffff","","","no")
    end if
'    xHTML = xHTML & "<br>"


ib=1
redim preserve aBloques(0)
redim preserve aBloquess(0)
redim preserve aBloquest(0)


'  Strsql = "SELECT * From Objtabblk Where Tabdef = " & nTabla & " order by tabblksec"
'  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'  do until rs.eof
'    nID = rs(0)

    nSec = 10
    nBTitulo = "BLOQUE 1"
    if len(nBTitulo) > 0 then
       ltitu = len(nBTitulo)
       if nBTitulo <> space(ltitu) and Ucase(nBTitulo) <> "NO" then

    redim preserve aBloques(ib)
    redim preserve aBloquess(ib)
    redim preserve aBloquest(ib)
    aBloques(ib) = 10
    aBloquess(ib) = 10
    aBloquest(ib) = "BLOQUE 1"
    ib = ib + 1

nStilo = " Style={font-family:" & chr(34) & nblkff & chr(34) & ";font-size:" & nblkfs & ";color:" & nblkfc & ";background-color:" & nblkbc & ";text-decoration:none;} "

nfonty = " <font face=" & chr(34) & replace(nblkff,"-"," ") & chr(34) & " color=" & nblkfc & " >"



      end if
    end if

    fila = ""
'    strsql = "SELECT * From Objtabcon Where tabblk = " & nID & " order by tabcontipo,tabconsec"
'    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

'    do until rs2.eof

nStilot = " Style={font-family:" & chr(34) & nctitff & chr(34) & ";font-size:" & nctitfs & ";color:" & nctitfc &  ";background-color:" & nctitbc & ";} "
nfontt = " <font face=" & chr(34) & replace(nctitff,"-"," ") & chr(34) & " color=" & nctitfc & " >"


nStiloc = " Style={font-family:" & chr(34) & ncconff & chr(34) & ";font-size:" & ncconfs & ";color:" & ncconfc &  ";background-color:" & ncconbc & ";} "
nfontc = " <font face=" & chr(34) & replace(ncconff,"-"," ") & chr(34) & " color=" & ncconfc & " >"

' titulos de columnas

      'Tipo

         nStilo = nStilot
         nFontr = nctitbc


      'Columnas
      celda = ""
      for ix = 0 to UBound(aColumnas)
        if cint(aColumnas(ix)) <> 0 then 
           sSize = cstr(aColumnas(ix)) & "% "
           sConten = "TIT.COLUMNA " & cstr(ix + 1)
           lconten = len(sconten) 
           select case cstr(aColumnasa(ix))
                  case "0" 
                         xalig = " align=center "
                  case "1" 
                         xalig = " align=left "
                  case "2" 
                         xalig = " align=right "
           end select
           if sConten = "" or sConten = space(lconten) then sConten = "&nbsp;"
           if len(sConten) = 0 then sConten = "<font color=" & nFontr & ">-</font>" 
           sConten = nfontt &  replace(sConten,chr(13),"<br>") & nfontx
           Celda = Celda & ofv.GenCelda("","aaa" & nStilo & xalig,"",sSize,"","","",sConten,"si")
        end if
      next

      Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

' Contenido de columnas

      'Tipo
      nStilo = nStiloc
      nFontr = ncconbc


      'Columnas
      celda = ""
      for ix = 0 to UBound(aColumnas)
        if cint(aColumnas(ix)) <> 0 then 
           sSize = cstr(aColumnas(ix)) & "% "
           sConten = "CONT.COLUMNA " & cstr(ix + 1)
           lconten = len(sconten) 
           if sConten = "" or sConten = space(lconten) then sConten = "&nbsp;"
           if len(sConten) = 0 then sConten = "<font color=" & nFontr & ">-</font>" 
           sConten =  nfontc &  replace(sConten,chr(13),"<br>") & nfontx
           select case cstr(aColumnasa(ix))
                  case "0" 
                         xalig = " align=center "
                  case "1" 
                         xalig = " align=left "
                  case "2" 
                         xalig = " align=right "
           end select
           Celda = Celda & ofv.GenCelda("","aaa" & nStilo & xalig,"",sSize,"","","",sConten,"si")
        end if
      next

      Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")



 '     rs2.movenext
 '   loop



'    call ofv.cerrarconsulta(rs2)

'    rs.movenext
'    loop

'  call ofv.cerrarconsulta(rs)

  call ofv.cerrarconn(cn)

'  if ubound(aBloques) > 1 then
'     response.write "hola2"

nStilo = " Style={font-family:" & chr(34) & nblkff & chr(34) & ";font-size:" & nblkfs & ";color:" & nblkfc & ";background-color:" & nblkbc & ";text-decoration:none;} "


    sHTML = sHTML & "<br>"
    Fila2 = ""
    for ib=1 to ubound(aBloques)
    npoint = "<a href=" & chr(34) & "#" & aBloques(ib) & chr(34) & nStilo & " >" & aBloquess(ib) & "&nbsp;&nbsp;" & aBloquest(ib) & "</a>"
    Celda = ofv.GenCelda("","","","100%","","","",npoint,"si")
    Fila2 =  ofv.GenRow("","","","","","","",Celda,"si")
    next

    sHTML = sHTML & ofv.GenTabla("","","","600","","0","0","0","0","#ffffff","",fila2,"si")
    sHTML = sHTML & "<br>"
      
'  end if
 
  sHTML = sHTML & xHTML & fila
'  sHTML = sHTML & ofv.GenfCelda()
'  sHTML = sHTML & ofv.GenfRow()
  sHTML = sHTML & ofv.GenfTabla()

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>