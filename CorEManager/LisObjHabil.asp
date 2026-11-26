<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
Server.ScriptTimeout=60*20
sub linea(xfila,xindi)
    if  xfila <> "" then
        response.write xfila
        if  indi > 10 then
            response.flush
            indi = 0
        end if
        xindi = xindi + 1
    end if
end sub

function Sqlgen(xcodobj,xtipobj)
dim Strsqlx


Strsqlx = "select * from csobjhabil "
Strsqlx = Strsqlx & "WHERE codobjmaster=" & xcodobj
Strsqlx = Strsqlx & " and tipobjmaster='" & xtipobj
Strsqlx = Strsqlx & "' and objun=" & session("UN")
Strsqlx = Strsqlx & " order by tipobjeto,"
Strsqlx = Strsqlx & " descri, codobjeto "

Sqlgen = Strsqlx
end function

function Objgen(xrs,xww,xtobj,xtobjd,xtest,xtestd)

dim wwx
dim wwy

xtx = "Esquema: "
for wwx=0 to Ubound(xtobj)
    if xtobj(wwx) = xrs(2) then
       xtx = xtobjd(wwx) & ": "
       exit for
    end if
next

wwy = "N"

if not isnull(xrs(5)) then
   if  xrs(6) or xrs(7) or xrs(8) or xrs(9) then
       wwy = "S"
       xindi = xindi + 1
   end if
end if

subx = "<tr class=cd HEIGHT=25 >"



subx = subx & "<td  class=VT >"
subx = subx & "<i class='las la-angle-right' style=color:" & Session("CMFPNTIT") & "></i></td>"

subx = subx & "<td  class=VTI >"
subx = subx & xtx & xrs(1) & "</td>"

xestad = "S/D"
for wwx=0 to Ubound(xtest)
    if cstr(xtest(wwx)) = cstr(xrs(3)) then
       xestad = xtestd(wwx)
       exit for
    end if
next


subx = subx & "<td  class=UT >"
subx = subx & cdate(clng(xrs(4)))
subx = subx & "</td>"

subx = subx & "<td  class=UT >"
subx = subx & xestad
subx = subx & "</td>"

subx = subx & "<td  class=UT >"
xxx=getinfouser( xrs(5) &"")
sOwner = session("SGU_nombre")
Owner = "&nbsp;&nbsp;" & xrs(5) & " - " & sOwner
subx = subx & Owner
subx = subx & "</td>"


if wwy = "N" then
   subx = subx & ""
else
    sInput = ofv.GenerarInput("boton1","Habilitar","submit","hab" & xrs(0),"5","bt","")
    sInput = ofv.gencelda("","ut","","5","","","",sInput,"si")
    sAccion = "edobjsdk.asp?Codobjeto=" & xrs(0) & "&Tipobjeto=" &  xrs(2) & "&Opcion=HE"
    subx = subx & ofv.GenForm(sForm & "1a","",sAccion,"","post","",sInput,"si")
end if

subx = subx & "</tr>"

if  wwy = "N" then subx = ""

ObjGen = subx

end function

function selecEmpre()

sImagenx = "<B><FONT FACE=CorFlatIcon COLOR=#aa0000 SIZE=4>x</FONT></B>"

   IF  Session("EMPMUN") = "S" THEN

   Fila    = ""

   SINPUT3 = "&nbsp;<A class=tt style='color:#000000;' title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

   Fila    = ""
   Celda   = ""
   sHref   = ""

   sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
   sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
   Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","",SINPUT3,"si")
   Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
   Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

   Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

   Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")





   shtml=shtml & "<script language=javascript>" & chr(13)
   shtml=shtml & "function shwDem(){" & chr(13)
   shtml=shtml & "document.all.DEM1.style.display='block';" & chr(13)
   shtml=shtml & "}" & chr(13)
   shtml=shtml & "</script>" & chr(13)
 '----------------------------------------------
   sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
   sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
   sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.all.DEM1.style.display='none';" & chr(34)
   sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
   TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:240px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


   TABLAX = TablaX & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
   TABLAX = TablaX & "<tr><td  "
   TABLAX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
   TABLAX = TablaX & "<tr><td colspan=3 "
   TABLAX = TablaX & " >&nbsp;</td></tr>"

   FOR IAEMP=1 TO TAEMP

    '''  RESPONSE.WRITE  ATREMP(IAEMP,2)

      IF ATREMP(IAEMP,2) = 1 THEN

       IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
          TABLAX = TablaX & "<tr BGCOLOR=#EEEEEE >"
       ELSE
          TABLAX = TablaX & "<tr>"
       END IF

       TABLAX = TablaX & "<td width=5 "
       IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
          TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
       ELSE
          TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
       END IF

       sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
       TABLAX = TablaX & "<td width=5 "
       TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

       TABLAX = TablaX & "<td class=vt  >"

       'sHref = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)
 		sHref = "LisObjHabil.asp?EMPRE=" & ATREMP(IAEMP,0)
     '''  sHref = "OBJtreeTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
       sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
       TABLAX = TablaX & sLink
       TABLAX = TablaX & "</td>"
       TABLAX = TablaX & "</tr>"
      END IF

   NEXT

   TABLAX = TablaX & OFV.genFtabla()

   TABLA = Tabla & TABLAX

   TABLA = Tabla & "</DIV>"

   sHTML = sHTML & TABLA

  end if

end function


ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0
  response.buffer = true
  CantRegMostrarx = 3
  CantRegMoverx = 1'cdbl(request.querystring("CantRegMover"))
    indi = 0
    sformh = "Habilitacion de Esquemas"

     XXXX = NEWEMPRE("NO")

	      session("ASP")="SUBCONSULTA.ASP"

	      sUN = request.querystring("EMPRE")
	    'response.write "sEMPRE: " & sEMPRE
	       IF sUN = ""  AND SESSION("UN") = "" THEN

	         sUN = ATREMP(0, 0)
	         SESSION("UN") = sUN & ""
	         If (SESSION("UN") & "") = "" Then SESSION("UN") = 1
	       elseIF sUN <> "" THEN
	        XXXX = NEWEMPRE(sUN)
	        SESSION("UN")=sUN
	       ELSE
	        XXXX = NEWEMPRE(SESSION("UN"))
	        sUN=SESSION("UN")
   end if

    shtml = ofv.FormHeader(sFormh)

    Celda =  ofv.GenCelda("","vt","center","","","","","<big><br><big>Cargando consulta...</big>un momento por favor</big>","si")
	 Fila  = ofv.GenRow("","","","","","","",celda,"si")

	 divx = "<div id='mt1' style='display:block;' >"
	 divx = divx & ofv.GenTabla("","","","60% style='border-width:2;border-color:#AAAAAA;' ","","0","","0","0","#ffffff","",Fila,"si")
 	 divx = divx & "</div></center>"

    TablaH = ofv.GenTabla("","","","20%","","0","0","0","25","","",selecEmpre(),"si")


	shtml = shtml & tablah

    response.write shtml
response.write divx

       response.write "<table class=aaa id=margen1 marginwidth=0 marginheight=0 style='visibility:hidden'; "
       response.write " cellspacing=0 cellpadding=2  border=0 >"
       response.write "<tr class=cd HEIGHT=25 >"
       response.write "<td width=1% class=UT >&nbsp;</td>"
       response.write "<td width=50% class=TTT >DOCUMENTO</td>"
       response.write "<td width=15% class=TTT >CREACION</td>"
       response.write "<td width=15% class=TTT >ESTADO</td>"
       response.write "<td width=15% class=TTT >PROPIETARIO</td>"
       response.write "<td width=5%  class=TTT >HABILITAR</td>"
       response.write "</tr>"



       response.flush

    set cn = ofv.conectar(ofv.strconn2)

'xx Arma tipos de objeto
dim tobj()
dim tobjd()
inx = 0
redim preserve tobj(0)
redim preserve tobjd(0)



Strsql = "Select * from tipobjeto order by tipobjeto"
set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
do until rs.eof
   redim preserve tobj(inx)
   redim preserve tobjd(inx)
   tobj(inx) = rs(0)
   tobjd(inx) = rs(1)
   inx = inx + 1
rs.movenext
loop

'xx Arma tipos de subtema
dim tsub()
dim tsubd()
inx = 0
redim preserve tsub(0)
redim preserve tsubd(0)

Strsql = "Select * from subtema order by tipobjeto"
set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
do until rs.eof
   redim preserve tsub(inx)
   redim preserve tsubd(inx)
   tsub(inx) = rs(0)
   tsubd(inx) = rs(1)
   inx = inx + 1
rs.movenext
loop
call ofv.cerrarconsulta(rs)

'xx Arma tipos de estado
dim test()
dim testd()
inx = 0
redim preserve test(0)
redim preserve testd(0)

Strsql = "Select * from estados order by objid"
set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
do until rs.eof
   redim preserve test(inx)
   redim preserve testd(inx)
   test(inx) = rs(0)
   testd(inx) = Ucase(rs(1))
   inx = inx + 1
rs.movenext
loop
aaa = 1
call ofv.cerrarconsulta(rs)
saltax = "<tr><td colspan=13  class=dt >&nbsp;</td></tr>"
'xx Arma Temas:
    Strsql = "Select * from tema order by descri"
    set rsa = ofv.crearconsultaEx(StrSql,cn,1,parametros)


    do until rsa.eof
       temax = "<tr><td colspan=13  class=vt >TEMA: " & rsa(1)  & "</td></tr>"
'xx Arma SubTemas:
       Strsql = "Select * from prodserv where codtema = " & rsa(0) & " order by descri"
       set rsb = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       do until rsb.eof
        tipox = "SUBTEMA: "
        for inx=0 to Ubound(tsub)
            if tsub(inx) = rsb(4) then
               tipox = tsubd(inx) & ": "
               exit for
            end if
        next
        subtemax = "<tr class=cd><td width=5 >&nbsp;</td>"
        subtemax = subtemax & "<td colspan=12  class=vt >" & tipox & rsb(1) & "</td></tr>"
'xx Arma Objetos C:
    set cnc = ofv.conectar(ofv.strconn2)
        Strsql = SqlGen(cstr(rsb(0)),cstr(rsb(4)))
        set rsc = ofv.crearconsultaEx(StrSql,cnc,1,parametros)
        if not rsc.eof then
       '    call linea(saltax,indi)
       '    call linea(temax,indi)
       '    call linea(subtemax,indi)
           temax = "<tr><td colspan=13  class=vt >&nbsp;</td></tr>"
        end if
        do until rsc.eof
           subtemax = Objgen(rsc,"2",tobj,tobjd,test,testd)
           call linea(subtemax,indi)
'xx Arma Objetos D:
    set cnd = ofv.conectar(ofv.strconn2)
        Strsql = SqlGen(cstr(rsc(0)),cstr(rsc(2)))
'        Strsql = SqlGen(rsc(0),rsc(2))
         set rsd = ofv.crearconsultaEx(StrSql,cnd,1,parametros)
         do until rsd.eof
           subtemax = Objgen(rsd,"3",tobj,tobjd,test,testd)
           call linea(subtemax,indi)
'xx Arma Objetos E:
    set cne = ofv.conectar(ofv.strconn2)
        Strsql = SqlGen(cstr(rsd(0)),cstr(rsd(2)))
'        Strsql = SqlGen(rsd(0),rsd(2))
          set rse = ofv.crearconsultaEx(StrSql,cne,1,parametros)
          do until rse.eof
             subtemax = Objgen(rse,"4",tobj,tobjd,test,testd)
             call linea(subtemax,indi)
'xx Arma Objetos F:
    set cnf = ofv.conectar(ofv.strconn2)
        Strsql = SqlGen(cstr(rse(0)),cstr(rse(2)))
'        Strsql = SqlGen(rse(0),rse(2))
           set rsf = ofv.crearconsultaEx(StrSql,cnf,1,parametros)
           do until rsf.eof
              subtemax = Objgen(rsf,"5",tobj,tobjd,test,testd)
              call linea(subtemax,indi)
'xx Arma Objetos G:
    set cng = ofv.conectar(ofv.strconn2)
        Strsql = SqlGen(cstr(rsf(0)),cstr(rsf(2)))
'        Strsql = SqlGen(rsf(0),rsf(2))
            set rsg = ofv.crearconsultaEx(StrSql,cng,1,parametros)
            do until rsg.eof
               subtemax = Objgen(rsg,"6",tobj,tobjd,test,testd)
               call linea(subtemax,indi)
'xx Arma Objetos H:
    set cnh = ofv.conectar(ofv.strconn2)
        Strsql = SqlGen(cstr(rsg(0)),cstr(rsg(2)))
'        Strsql = SqlGen(rsg(0),rsg(2))
             set rsh = ofv.crearconsultaEx(StrSql,cnh,1,parametros)
             do until rsh.eof
                subtemax = Objgen(rsh,"7",tobj,tobjd,test,testd)
                call linea(subtemax,indi)
'xx Arma Objetos I:
    set cni = ofv.conectar(ofv.strconn2)
        Strsql = SqlGen(cstr(rsh(0)),cstr(rsh(2)))
'        Strsql = SqlGen(rsh(0),rsh(2))
              set rsi = ofv.crearconsultaEx(StrSql,cni,1,parametros)
              do until rsi.eof
                 subtemax = Objgen(rsi,"8",tobj,tobjd,test,testd)
                 call linea(subtemax,indi)
'xx Arma Objetos J:
    set cnj = ofv.conectar(ofv.strconn2)
        Strsql = SqlGen(cstr(rsi(0)),cstr(rsi(2)))
'        Strsql = SqlGen(rsi(0),rsi(2))
               set rsj = ofv.crearconsultaEx(StrSql,cnj,1,parametros)
               do until rsj.eof
                  subtemax = Objgen(rsj,"9",tobj,tobjd,test,testd)
                  call linea(subtemax,indi)
'xx Arma Objetos K:
    set cnk = ofv.conectar(ofv.strconn2)
        Strsql = SqlGen(cstr(rsj(0)),cstr(rsj(2)))
'        Strsql = SqlGen(rsj(0),rsj(2))
                set rsk = ofv.crearconsultaEx(StrSql,cnk,1,parametros)
                do until rsk.eof
                   subtemax = Objgen(rsk,"10",tobj,tobjd,test,testd)
                   call linea(subtemax,indi)
'xx Arma Objetos L:
    set cnl = ofv.conectar(ofv.strconn2)
        Strsql = SqlGen(cstr(rsk(0)),cstr(rsk(2)))
'        Strsql = SqlGen(rsk(0),rsk(2))
                 set rsl = ofv.crearconsultaEx(StrSql,cnl,1,parametros)
                 do until rsl.eof
                    subtemax = Objgen(rsl,"11",tobj,tobjd,test,testd)
                    call linea(subtemax,indi)

                 rsl.movenext
                 loop
                 call ofv.cerrarconsulta(rsl)
     call ofv.cerrarconn(cnl)
                rsk.movenext
                loop
                call ofv.cerrarconsulta(rsk)
     call ofv.cerrarconn(cnk)
               rsj.movenext
               loop
               call ofv.cerrarconsulta(rsj)
     call ofv.cerrarconn(cnj)
              rsi.movenext
              loop
              call ofv.cerrarconsulta(rsi)
     call ofv.cerrarconn(cni)
             rsh.movenext
             loop
             call ofv.cerrarconsulta(rsh)
     call ofv.cerrarconn(cnh)
            rsg.movenext
            loop
            call ofv.cerrarconsulta(rsg)
     call ofv.cerrarconn(cng)
           rsf.movenext
           loop
           call ofv.cerrarconsulta(rsf)
     call ofv.cerrarconn(cnf)
          rse.movenext
          loop
          call ofv.cerrarconsulta(rse)
     call ofv.cerrarconn(cne)
         rsd.movenext
         loop
         call ofv.cerrarconsulta(rsd)
    call ofv.cerrarconn(cnd)
        rsc.movenext
        loop
        call ofv.cerrarconsulta(rsc)
    call ofv.cerrarconn(cnc)
       rsb.movenext
       loop
       call ofv.cerrarconsulta(rsb)
    rsa.movenext

    loop


  sHTML =  "</table>"

  if  indi = 0 then
      sHTML = sHTML & "<br><br><p >&nbsp;&nbsp;<b><font face=""corflaticon"" size=4 color=""#993333"" >&#0105;</font></b>"
      sHTML = sHTML & "&nbsp;&nbsp;<font size=3 > No se registran documentos restringidos o inhabilitados.</font></p> "
  end if
   ' sHTML = sHTML & ofv.botoneraESSX("2",rsa,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda & celda2)

   shtml = "<script language=javascript >"
   shtml = shtml & "document.getElementById('mt1').style.display='none';"
   shtml = shtml & "document.getElementById('mt1').style.visibility='hidden';"
   shtml = shtml & "document.getElementById('margen1').style.visibility='';"
   shtml = shtml & "</script>"

  sHTML = sHTML & "<br><br></body></html>"
end if
call ofv.cerrarconsulta(rsa)
call ofv.cerrarconn(cn)
response.write sHTML
response.flush
%>
