<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  bModificar = False
  ofv.ObtenerAtributos Session("Form"),""
  bModificar = ofv.AModificar

  set cn = ofv.conectar(ofv.strconn6)

''background-color:" & session("CMFMNBOT") & ";background-color:" & session("CMFMENU") & ";

    sImagenP = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>&#0143;</FONT></B>"
    sImagenR = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>2</FONT></B>"  

stilo1 = " td.a1 {border-style:ridge;border-width:1;border-color:#f7aa22;position:relative;top:+5;}"
stilo2 = " td.a2 {border-color:#f7aa22;border-top-style:inset;border-left-style:inset;border-right-style:inset;border-width:1;position:relative;top:+2;}"
stilo3 = " td.a3 {border-color:#f7aa22;border-bottom-style:inset;border-left-style:inset;border-right-style:inset;border-width:1;}"
stilo4 = " td.a4 {background-color:#f7aa22;border-style:ridge;border-width:1;border-color:#f7aa22;position:relative;top:+5;}"



'xxxxxxxxxxxxxxxxxx
  StrSql = "Select * From REQUERIMIENTO WHERE CODIGO = '" & SESSION("NCLIENTE") & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        TITULO = rs(2)
        ASIGNA = rs(14)
        treq   = rs(15)
        cinter = rs(19) & ""
     end if

dim atab(15,1)
dim atabN(15,1)


    for xx = 1 to 15
           atab(xx,0) = 99
           atabN(xx,0) = 99
    next

    for xx = 1 to 9

           atab(XX,0) = xx + 2



    next



     StrSql = "Select * From TIPOREQ WHERE TIPOREQ = '" & treq & "' "
     set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rsX.eof then




        for yy = 1 to 15
 
            if clng(atab(yy,0)) = 99 then exit for
               xx = atab(yy,0)

               IF RSX(XX) THEN
                  atabN(yy,0) = XX
                  atabN(yy,1) = rsX(xx + 15)
               ELSE
                  atabN(yy,0) = 99
               END IF 
        next

     end if




  sHTML = "<html>"
  sHTML = sHTML & ofv.FormStyler()

 ' if sFuncion = "E" then

 ' else
 '    sHTML = sHTML & "<body bgcolor=#ffffff  >"
 ' end if


     sHTML = sHTML & "<body  onload='document.init.submit();' >"

  sHTML = sHTML & "<style>" & stilo1 & stilo2 & stilo3 & stilo4 & "</style>"

  T1 =  ofv.GenTabla("panel","xxx ","","100% border=0 ","",0,0,-1,-2,"","","","no")


    bstitulo = " REQUERIMIENTO : <b>" & TITULO & "</b>"
    bstitulo = "&nbsp;" & Ucase(bstitulo)

    h1 = "<IMG  src='Imagenes/ICorAbierto.jpg'  border=0 >"
    h1 = ofv.GenCelda("","VT","","3%","","","",h1,"si")
    h1 = h1 & ofv.GenCelda("","VT","","","","","",bstitulo,"si")

    if cinter = "" then cinter = session("NCliente")

    IF ASIGNA = SESSION("USUARIO") THEN

       StrSql = "Select * From PROYECTO WHERE CODIGO = '" & SESSION("NCLIENTE") & "' "
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rs.eof then
          t3 = "<a class=cc href='BARRAPRY.asp"
          t3 = t3 & "' <B><U>"
          t3 = t3 & "VER PROYECTO ASOCIADO&nbsp;" & sImagenP
          t3 = t3 & "</U></B></A>"
         h1 = h1 & ofv.GenCelda("","VT","RIGHT","30%","","","",t3,"si")
       ELSE 
         h1 = h1 & ofv.GenCelda("","VT","CENTER","30%","","","","<b>CODIGO:&nbsp;" & cinter & "</b>","si")
       end if
    ELSE 
         h1 = h1 & ofv.GenCelda("","VT","CENTER","30%","","","","<b>CODIGO:&nbsp;" & cinter & "</b>","si")
    end if




  T1 = T1 & ofv.GenRow("","","","","","","",h1,"si")
  T1 = T1 & ofv.GenFTabla()


titc = 0
 
    T2 =  ofv.GenTabla("xt","xxx ","","100%","",0,0,-1,0,"","","","no")
    T2 = T2 & ofv.GenRow("","","","","","","","","no")

            clase = "a1"

  if session("ModMain") = "WK" then

      sHTML = sHTML & "<form name=init id=init action='CoWKFREQUER.asp' target='Principal' ></form>"

      t3 = "<a class=cc href='CoWKFREQUER.asp"
      t3 = t3 & "' target=Principal><B>"
      t3 = t3 & "&nbsp;WORKFLOW&nbsp;"
      t3 = t3 & "</B></A>"
      T2 = T2 & ofv.GenCelda("","a2  id=xs" & titc & " name=" & clase & "  onclick='resalta(this);' ","center","","","","",t3,"si")

      titc = titc + 1
      t3 = "<a class=cc href='CoREQUER.asp"
      t3 = t3 & "' target=Principal><B>"
      t3 = t3 & "&nbsp;REQUERIMIENTO&nbsp;"
      t3 = t3 & "</B></A>"
      T2 = T2 & ofv.GenCelda("","a1  id=xs" & titc & " name=" & clase & "  onclick='resalta(this);' ","center","","","","",t3,"si")


  else

      sHTML = sHTML & "<form name=init id=init action='CoREQUER.asp' target='Principal' ></form>"

      t3 = "<a class=cc href='CoREQUER.asp"
      t3 = t3 & "' target=Principal><B>"
      t3 = t3 & "&nbsp;REQUERIMIENTO&nbsp;"
      t3 = t3 & "</B></A>"
      T2 = T2 & ofv.GenCelda("","a2  id=xs" & titc & " name=" & clase & "  onclick='resalta(this);' ","center","","","","",t3,"si")

  end if



        for yy = 1 to 15
 
            if clng(atabN(yy,0)) < 99 then
            titc = titc + 1

            clase = "a1"

            StrSql = "Select * From reqinfo WHERE CODIGO = '" & SESSION("NCLIENTE") & "' and tipo = " & atabN(yy,0)
            set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not rsX.eof then
               clase = "a4"
            end if

            t3 = "<a class=cc href='CoSolapa.asp?Solapa=" & atabN(yy,0) & "&NSOLAPA=" & replace(Ucase(atabN(yy,1))," ","&nbsp;")
            t3 = t3 & "' target=Principal><B>"
            t3 = t3 & "&nbsp;" & replace(Ucase(atabN(yy,1))," ","&nbsp;") & "&nbsp;"
            t3 = t3 & "</B></A>"

            T2 = T2 & ofv.GenCelda("",clase & " id=xs" & titc & " name=" & clase & " onclick='resalta(this);' ","center","","","","",t3,"si")
            end if


        next






  T2 = T2 & ofv.GenFRow()
  T2 = T2 & ofv.GenRow("","","","","","","","","no")

'''  if titc = 0 then titc = 1

  T2 = T2 &  "<td class=a3 colspan=" & titc + 1  & "  ><hr size=1 style='border-style:ridge;' color=#f7aa22></TD>"

  T2 = T2 & ofv.GenFRow()
  T2 = T2 & ofv.GenFTabla()

  sHTML = sHTML & "<script language=javascript>"
  sHTML = sHTML & " function resalta(yo) { "
  sHTML = sHTML & " for (i=0;i<" & titc + 1 & ";i++)"
  sHTML = sHTML & " {"
  sHTML = sHTML & " document.all('xs' + i).className = document.all('xs' + i).name ; "
  sHTML = sHTML & " } "
  sHTML = sHTML & " document.all(yo.id).className = 'a2'; "
  sHTML = sHTML & " } "
  sHTML = sHTML & "</script>"


    celda = ofv.GenCelda("","vt","","100%","","","",t1 & t2,"si")
    celda = ofv.GenRow("","","","","","","",celda,"si")
    sHTML = sHTML & ofv.GenTabla("","aaa style='filter:shadow(color=#dddddd,direction=125);' ","","100%","","0","","-2","-1","","",celda,"si")


  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>
