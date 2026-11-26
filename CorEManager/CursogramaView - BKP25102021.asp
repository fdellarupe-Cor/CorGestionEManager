<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else


'response.write Request.querystring

    sOCesque = Request.querystring("Codobjeto")
    nCodigo = Request.querystring("Codproced")
    sTipo = Request.querystring("Tipobjeto")
    vuelta = Request.querystring("Volver")
    sOTesque=sTipo
    Set cn = OFV.conectar(OFV.strconn2)


    StrSql = "select IDSEC  from OBJPROCURSODET where codobjeto = " & sOCesque
  '''  response.write strsql
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If rs.EOF Then
       response.redirect "Corvertodoseventos.asp?Codproced=" & nCodigo & "&Codobjeto=" & sOCesque & "&Tipobjeto=" & sTipo
    End If
    Call OFV.cerrarconsulta(rs)

    If InStr(Request.ServerVariables("HTTP_User_Agent"), "MSIE 6") Then 'IE6
		    ie="N"
		ElseIf InStr(Request.ServerVariables("HTTP_User_Agent"), "MSIE 8") Then 'IE7
		   ie="N"
		 else
		 ie="S"
   end if


%>
<!DOCTYPE html>
<html>
   <head>



     <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>


     <script type="text/javascript" src="JSGRAF/js-graph-it.js"></script>


     <link rel="stylesheet" type="text/css" href="JSGRAF/js-graph-it.css">
     <link rel="stylesheet" type="text/css" href="JSGRAF/sf-homepage.css" />
     <link rel="stylesoheet" href="JSGRAF/demo-print.css" type="text/css" media="print">
 <style>

      /* Apply the element you want to drag/resize */
 	 .drsElement {
 	  position: absolute;
 	  border: 1px solid #333;
 	 }

          .rotateDiv	{
 	   		   			margin-top:300px;
 	   		   			-webkit-transform:rotate(120deg);
 	   		   			-moz-transform:rotate(120deg);
 	   		   			-o-transform:rotate(120deg);
 	   		   			/* filter:progid:DXImageTransform.Microsoft.BasicImage(rotation=1.5); */
 	   		   			-ms-transform:rotate(270deg);

 		   		}

        .block { position: absolute; }


 .dd1{

                      border-left-style:solid;
                      border-left-width:1px;
                      border-left-color:#99ccff;
                      border-bottom-style:solid;
                      border-bottom-width:1px;
                      border-bottom-color:#99ccff;
                      font-size:8px;
 		     font-face:arial;
 }


 .dd{
         border: 1px SOLID #99ccff;
                      font-size:8px;
 		     font-face:arial;
 }


        .container_2 {

                      font-size:10px;
 		     font-face:arial;
 		     background: #f7eedd;



 }

 .dv{
         border: none; }


        #container_4 {

                      font-size:12px;
 		     font-face:arial;


 }


 .dx{
         border: 1px solid #eeeeee; }


        #container_5 {

                      font-size:10px;
 		     font-face:arial;
 		     background: #F7BE81;

                      }

   .tttt
   {
   position: absolute;
   z-index:1000;
   }

   .dhtmltooltip{
 position: absolute;
 left: 00px;
 width: auto;
 qheight:100px;
 border: 1px solid black;
 padding: 2px;
 background-color: lightyellow;
 font-family: arial;
  font-size:9pt;
 font-style:italic;
 visibility: hidden;
 z-index: 1000;

 filter: progid:DXImageTransform.Microsoft.Shadow(color=gray,direction=135);
 }

 #dhtmlpointer{
 position:absolute;
 left: -300px;
 z-index: 101;
 visibility: hidden;

 }



     </style>
   </head>



<%
''   Dim cn
''   dim sconectores
 ''  set cn=Server.CreateObject("ADODB.Connection")


    Set cn = OFV.conectar(OFV.strconn2)

   bstitulox = "none............"


    StrSql = "Select B.descri, A.codiden, A.descri, a.* From OBJetos A inner join tipobjeto B on (A.tipobjeto = b.tipobjeto) Where a.codobjeto  = " & sOCesque
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    objun = "1"
    If Not rs.EOF Then
	objun = rs("objun")
      if session("FULLDOCN") = "1" then
         bstitulox = Capitalize(rs(0) & "") & ": " & Capitalize(rs(2))
      elseif session("FULLDOCN") = "2" then
         bstitulox = Capitalize(rs(0) & "") & ": " & rs(1) & " - " & Capitalize(rs(2))
      else
         bstitulox = Capitalize(rs(0) & "") & ": " & rs(1) & " - " & Capitalize(rs(2))
      end if


    End If
    Call OFV.cerrarconsulta(rs)
 if objun=1 and Session("EMPMUN") = "S" then
 	  sql="exec SpSectoresCorp " & sOCesque &",'" & sTipo & "'," & session("un")
 	'     response.write sql
 	  Set rsSecCorp = ofv.crearconsultaEx(sql,cn,1,parametros)

 		 if not rsSecCorp.eof then

 			 ASecCorp =rsSecCorp.getrows
 		  end if
 		Call ofv.cerrarconsulta(rsSecCorp)
  end if

   muestra=request.querystring("Muestra")
   'conn.Provider="Provider=SQLOLEDB;Data Source=server-20;Initial Catalog=CDIChat;User Id=sa;Password=bb4515;"
''   cn.Open "Provider=SQLOLEDB;Data Source=server-20;Initial Catalog=CDIChat;User Id=sa;Password=bb4515;"

   sconectores=""
  DIM APRCEV(2000,3)
  DIM ASEC(2000,3)
  DIM ASEC2(2000,2)
    IALINKC=0
   I=1
   ISECPOS=50
   isec = 0
   sidcon=""

  '' set rs= cn.execute("select DISTINCT IDSEC  from libro ")

  StrSqlx = "select DISTINCT IDSEC  from OBJPROCURSODET where codobjeto = " & sOCesque
  Set rs = OFV.crearconsultaEx(StrSqlx, cn, 1, "")

   DO WHILE  NOT RS.EOF
   'ASEC(I,0)=RS(0)
   ASEC2(I,1)=ISECPOS
   ASEC2(I,0)=RS(0)
   isec = i
   I=I+1
   ISECPOS=ISECPOS + 300
   RS.MOVENEXT
LOOP

   I=1
   iprc = 0

  StrSqlx = "select max(codproced),  max(id)  as iden "
  StrSqlx = Strsqlx & " from OBJPROCURSODET "
  StrSqlx = Strsqlx & " where codobjeto = " & sOCesque
  StrSqlx = Strsqlx & " group by codproced "
  StrSqlx = Strsqlx & " order by iden "

'''  StrSqlx = "select DISTINCT codproced  from OBJPROCURSODET where codobjeto = " & sOCesque
  Set rs = OFV.crearconsultaEx(StrSqlx, cn, 1, "")

   DO WHILE  NOT RS.EOF

      APRCEV(I,0)=RS(0)

      StrSqlx = "select b.evendescri, b.perfdescri,b.secuen,b.secuencia from CSCEventoSectorPerfil b where b.codproced = '" & rs(0) & "' "
      Set rs2 = OFV.crearconsultaEx(StrSqlx, cn, 1, "")
      if rs2.eof then
         APRCEV(I,1)=""
         APRCEV(I,2)=""
      else
          APRCEV(I,1)=rs2(2) & "." & rs2(3)  &" " & RS2(0)
         APRCEV(I,2)=RS2(1)
      end if


      iprc = i
      I    = I+1

   RS.MOVENEXT
   LOOP




  StrSqlx = "select * FROM OBJPROCURSOCMP where codobjeto = " & sOCesque
  Set rs2 = OFV.crearconsultaEx(StrSqlx, cn, 1, "")

  IF  NOT RS2.EOF THEN

      WWWNONPDF = RS2(4) & ""

  ELSE

      WWWNONPDF = "NO"
  END IF



  StrSqlx = " select A.*, b.idsec as SECSI, c.idsec as SECNO, b.id as SECSIid , c.id as SECNOid,x.archivocur, d.idsec as SECCN,D.id as SECCNid "
  StrSqlx = StrSqlx &  " from OBJPROCURSODET A "
   StrSqlx = StrSqlx &  " left join modo x on(x.tipomodo COLLATE DATABASE_DEFAULT=a.modo COLLATE DATABASE_DEFAULT) "
  StrSqlx = StrSqlx &  " left join OBJPROCURSODET b on (A.codaccionSI=b.codaccion and A.codobjeto = b.codobjeto) "
  StrSqlx = StrSqlx &  " left join OBJPROCURSODET c on (A.codaccionNO=c.codaccion and A.codobjeto = c.codobjeto) "
   StrSqlx = StrSqlx &  " left join OBJPROCURSODET d on (A.codaccioncn=d.codaccion and A.codobjeto = d.codobjeto) "
  StrSqlx = StrSqlx &  " where a.codobjeto = " & sOCesque  & " "
  StrSqlx = StrSqlx &  " ORDER BY A.ID "

'response.write strsqlx

  Set rs = OFV.crearconsultaEx(StrSqlx, cn, 1, "")


%>
   <body   style='backgrqound-color: #e5e5e5' onload='poselm();xposConectores();'>






 <script language=JavaScript>
         function AbrirCons(Datos,nom) {
              myWin= open(Datos,"ConsLnk" + nom,"menubar=no,toolbar=no,height=300,width=600,scrollbars=yes,resizable=yes");
           }
         function pointcon(xcon) {
              window.scrollTo(document.all(xcon).offsetParent.offsetLeft,document.all(xcon).offsetParent.offsetParent.offsetParent.offsetTop + document.all(xcon).offsetParent.offsetTop);
           }
 </script>


<%


  StrSqlx = " select distinct a.sector,a.IDSEC,cast(b.secuen as nvarchar) + '.'+ cast(b.secuencia as nvarchar) +  ' ' + b.evendescri as evento,A.IDSECTOR  from OBJPROCURSODET  a "
    StrSqlx = " select distinct a.sector,a.IDSEC,A.IDSECTOR from OBJPROCURSODET  a "
'strsqlx=strsqlx & "inner join CSCEventoSectorPerfil b on (b.codproced=a.codproced)"
  strsqlx=strsqlx & " where a.codobjeto = " & sOCesque & " order by a.idsec "
'response.write strsqlx
''''response.write strsqlx

  Set rss = OFV.crearconsultaEx(StrSqlx, cn, 1, "")


'''''''''''''''set rss= cn.execute("select distinct sector,IDSEC from OBJPROCURSODET  where codobjeto = " & sOCesque & " order by idsec " )
IST=1
    DO WHILE  NOT RSs.EOF
       if isarray(ASecCorp) and objun=1 and Session("EMPMUN") = "S" then
	   	   	 		SECD  = UCASE(RSS("SECTOR"))
	   	   	 		SECD = FindInArray(ASecCorp,6,rsS("IDSECTOR"),1,SECD)

	   	   	 	else
	   	   	 		SECD  = UCASE(RSS("SECTOR"))
       end if
       RESPONSE.WRITE  " <div  id=ECO" & IST & " class=dhtmltooltip onclick=""this.style.visibility ='hidden';"">" & SECD & "</div>"


       RSS.MOVENEXT
       IST=IST+1
    LOOP
    RSS.CLOSE
%>


     <script type="text/javascript">



  function ver(t)
     {
      var zz;
      zz='<html><head><link rel="stylesheet" type="text/css" href="JSGRAF/js-graph-it.css"><link rel="stylesheet" type="text/css" href="JSGRAF/sf-homepage.css" /><head>'

      zz= zz + document.getElementById('mainCanvas').innerHTML;

 //zz= document.getElementById('mainTabla').innerHTML;
    t.fname.value=zz;
 //   alert(document.getElementById('mainCanvas').style.width);
//alert(document.getElementById('mainCanvas').style.posHeight);
   // t.largo.value=document.getElementById('mainTabla').clientWidth;
  //  t.alto.value= document.getElementById('mainTabla').clientHeight;
     }
function ver2(t)
     {

     var dato='';
     var zz= document.getElementById('mainCanvas');

     var ele=getElementsByClassName(zz,'block');

     for(i=0;i< ele.length;i++)
     {
      posx=document.getElementById (ele[i].id).style.left;
      posy=document.getElementById (ele[i].id).style.top;
      dato=dato + ele[i].id + '###'  +   posx + '###' + posy + '///'
     }

    t.xfname.value=dato;
     }



function getElementsByClassName(node, classname) {
    var a = [];
    var re = new RegExp('(^| )'+classname+'( |$)');
    var els = node.getElementsByTagName("*");
    for(var i=0,j=els.length; i<j; i++)
        if(re.test(els[i].className))a.push(els[i]);
    return a;
}



function checkS(e,id,mod){


    if (mod=='1') {
        var _x;
	var _y;




		if (e.pageX >0 || e.pageY > 0) 	{
			_x = e.pageX;
			_y = e.pageY;

		}
		else if (e.clientX || e.clientY) 	{
			_x = event.clientX + document.documentElement.scrollLeft;

			_y =event.clientY + document.documentElement.scrollTop;

	}

        posx = _x;
	posy = _y;


<%


        FOR TI=1 TO IST-1
            RESPONSE.WRITE "  document.getElementById('ECO"&TI&"').style.visibility = ""hidden"";"
        NEXT


%>

       document.getElementById(id).style.left = _x + 'px';;
       document.getElementById(id).style.top = _y + 'px';;

       document.getElementById(id).style.visibility = "visible";




                  }

       else

                  {

       document.getElementById(id).style.visibility = "hidden";

                  }


}

       </script>







      <div class="canvas" id="mainCanvas"
       style="overflow:hidden;visibility:hidden;backgrowund-color: #e5ew5e5;width:100%;height:600PX">

      <%
       xtabla = "<table  id='mainTabla'  width=500px height=0px cellpadding=0 cellspacing=0 style='border-style:solid;border-width:1px;border-color:#99ccff;backwground-color: #e5we5e5'>"
      'ssql = "select distinct sector,idsec,max(posy)-min(posy) AS ALTO from OBJPROCURSODET where codobjeto = " & sOCesque & " group by sector,idsec order by idsec "
       ssql = "select distinct a.sector,a.idsec,max(posy)-min(posy) AS ALTO,c1.cliente,A.IDSECTOR "
	   ssql = ssql & " from OBJPROCURSODET a "
	   ssql = ssql & " inner join csxobjsector b on (a.codobjeto=b.codobjeto) "
	   ssql = ssql & " inner join eco c on (c.eco=b.codsector and c.descri=a.sector) "
	   ssql = ssql & " inner join clienteplan c1 on (c.ecoUN=c1.Id) "
	   ssql = ssql & " where a.codobjeto =" & sOCesque
	   ssql = ssql & " group by a.sector,a.idsec,c1.cliente,A.IDSECTOR "
	   ssql = ssql & " order by a.idsec "
	   'response.write ssql
	   Set rss = OFV.crearconsultaEx(ssql,cn,1,parametros)

       xtabt = 0

	     xtablax = ""

	     xtablax = xtablax & "<tr >"

             xtablax = xtablax & " <td class=dd1 width=2px height=20px align=left>&nbsp;</td>" & CHR(13)

	     for x=1 to iprc
	         xtablax = xtablax & " <td width=200px  class=dd1 height=20px  id='main" & APRCEV(x,0) & "' align=left>" & CHR(13)
                 xtablax = xtablax & "&nbsp;" & ucase(APRCEV(x,1) & "") & "-" & ucase(APRCEV(x,2) & "")
                 xtablax = xtablax & "</td>" & CHR(13)
                 xtabt   = x
	     next

             xtablax = xtablax &" </tr>"


          wbgcolor="xx"
          ist=0

	  xtabla = xtabla & "<tr >"


	  xtabla = xtabla & " <td colspan=" & (1 + xtabt) & " id=container_4  class=dv height=20px style='border-bottom-style:solid;border-bottom-width:1px;border-bottom-color:#99ccff;' align=left>" & CHR(13)
          xtabla = xtabla & "&nbsp;" & ucase(bstitulox)

          IF WWWNONPDF <> "NO" THEN
             xtabla = xtabla & "&nbsp;-&nbsp;<a href='PnlPrint/" & WWWNONPDF & "' TARGET='_blank' ><u>VER PDF</u></a>"
          END IF




          xtabla = xtabla & "</td>" & CHR(13)
          xtabla = xtabla & " </tr>"

          xtabla = xtabla & xtablax

	     DO WHILE  NOT RSs.EOF

                xsec=""
                ist = ist + 1


           if wbgcolor="" then
              wbgcolor=" bgcolor='#eeeeee' "
           else
              wbgcolor=""
           end if

	   xtabla = xtabla & "<tr >"

          xalto =  300
          if RSS("ALTO") > xalto  then
             xalto =  RSS("ALTO")
          end if

	   xtabla = xtabla & " <td  class=container_2 width=2px height=" & xalto &"px align=left>" & CHR(13)
	          SSECTOR=REPLACE(rss("sector")," ","")
	          SSECTOR=REPLACE(rss("sector"),CHR(9),"")

	          if isarray(ASecCorp) and objun=1 and Session("EMPMUN") = "S" then

			   SSECTOR = FindInArray(ASecCorp,6,rsS("IDSECTOR"),1,SSECTOR)
			   SSECTOR=REPLACE(SSECTOR," ","")
			   SSECTOR=REPLACE(SSECTOR,CHR(9),"")

              end if

	  	     for x=1 to len( left(SSECTOR,20))'
	  	     xsec=xsec &  mid(SSECTOR,x,1) & "<br>"
	  	     next
	  	     if session("EMPMUN") = "S" and objun = 1 then
			 	'xtabla = xtabla  & "<img   id=ISEC" & ist & " class=secox title='" & rss("cliente") & "&nbsp;-&nbsp;" & left(SSECTOR&"",40) & "' style='display:none'></img><span id=span"& ist & " class=spanxx>" & ucase(xsec)&"</span>"
			 	 xtabla = xtabla  & "<img   id=ISEC" & ist & " class=secox title='" &  "&nbsp;-&nbsp;" &  left(SSECTOR&"",40) & "' style='display:none'></img><span id=span"& ist & " class=spanxx>" & ucase(xsec)&"</span>"
			 else
		     	xtabla = xtabla  & "<img   id=ISEC" & ist & " class=secox title='"  &  left(SSECTOR &"",40) & "' style='display:none'></img><span id=span"& ist & " class=spanxx>" & ucase(xsec)&"</span>"
		     end if

	  xtabla = xtabla & "</td>" & CHR(13)

	'  xtabla = xtabla &" <td width=3200px height=200px  class=dd onclick=""checkS(event,'ECO" & ist & "','1')"" >"
	'  xtabla = xtabla & "&nbsp;"
	'  xtabla = xtabla &"  </td>"

	   for x=1 to iprc
	       xtabla = xtabla & " <td width=100px  class=dd1 height=" & xalto &"px onclick=""checkS(event,'ECO" & ist & "','1')"" align=left>&nbsp;</td>" & CHR(13)

	   next



          xtabla = xtabla &" </tr>"


    rss.movenext
    loop
     xtabla = xtabla &" </table>"
    response.write xtabla
     %>
      <%
      SHTML=""
      SHTMLx=""
      xpos=40
      ypos=50
      isect=rs("idsec")
      i=0

      iiprc = 0
      iprcant = ""

      xposant = 0
      wmodal = "P"
  '    wmodal = "E"

''''APRCEV(x,1)

      do while not rs.eof' and i <15
         xcols = rs("CODACCION")
         ENTRO="NO"
         Wtacc = "M"
         wnewprc = "N"

         if  rs("CODPROCED") <> iprcant THEN
             IF IIPRC > 0 THEN
             '''   APRCEV(IIPRC,3) = xposPREV
                APRCEV(IIPRC,3) = xposant
                wnewprc = "S"
                if   IIPRC = 1 then
                     APRCEV(IIPRC,3) = APRCEV(IIPRC,3) + 60
                else
                     APRCEV(IIPRC,3) = xposant
                end if
             END IF
             iprcant = rs("CODPROCED")
             IIPRC=1
             FOR III=1 TO IPRC
                 IF APRCEV(III,0) = iprcant THEN
                    IIPRC=III
                    EXIT FOR
                 END IF
             NEXT

         END IF


              if  wnewprc = "S" then
	          xpos     = xpos + 40
              end if



         if UCASE(left(xcols,3) & "") = "ALT" THEN
            Wtacc = "A"
         END IF
		ASEC(I,0) = xcols

		' if rs("idsec") <> isect then


                 For isecb = 1 To isec
                    If ASEC2(isecb,0) = rs("idsec") Then
		       ypos  = ASEC2(isecb,1)
                       IF  Wtacc = "A" THEN ypos  = ypos + 100
                       Exit For
                    End If
                 Next




                 IF Wtacc = "M" AND ucase(rs(3)&"") = "INICIO" THEN
	          '  xposPREV = xpos + 80
	            xposPREV = xpos + 50
                 else
                    wonka = split(rs(1)," ")
                    wonkk = ""
                    wonko = ""
                    wonkt = ""
                    wideset = "N"


                    for iii=0 to ubound(wonka)
                        wonko = wonko & wonkk & wonka(iii)

                        if wonkk = "&nbsp;" then
                           wonkk = "<BR>"
                        else
                           wonkk = "&nbsp;"
                        end if
                    next


                     '''      if len(wonkt) > 19 then wideset = "W"

                    if iii=2 then
                  ''     wonkt =  replace(wonko,"&nbsp;"," ")
                       wonko = replace(wonko,"&nbsp;","<BR>")

                   '''    if len(wonkt) > 19 then wideset = "W"
                    end if

                    wonka2 = split(replace(wonko,"&nbsp;"," "),"<BR>")



                if  wmodal = "E" then

                   lant = 0

                    for iii=0 to ubound(wonka2)

                        if len(wonka2(iii)) > lant then
                           lant = len(wonka2(iii))

                        if len(wonka2(iii)) < 11 then
                                    wideset = "N"
                           elseif len(wonka2(iii)) < 16 then
                                    wideset = "M"
                               else
                                    wideset = "W"
                               end if

                        end if
                    next



                    if  wideset = "W" then
                        xposPREV = xpos + 180
                    elseif  wideset = "M" then
                        xposPREV = xpos + 120
                    else
                        xposPREV = xpos + 100
                    end if

               else

                    wmodlen = 80

                    lant = 8

                    for iii=0 to ubound(wonka2)

                        if len(wonka2(iii)) > lant then
                           lant = len(wonka2(iii))

                           wmodlen = (len(wonka2(iii)) * 10)

                        end if
                    next

                    xposPREV = xpos + wmodlen

               end if


                    IF Wtacc = "M" AND ucase(rs(3)&"") = "DECISIÓN" THEN
                       xposPREV = xposPREV + 20
	               xpos     = xpos + 60
                    END IF

                  END IF

              '      IF Wtacc = "M" AND ucase(rs(3)&"") = "DECISIÓN" THEN
	       '       ' xposPREV = xpos + 180
	      '        xposPREV = xpos + 120
	       '        xpos     = xpos + 60
               '     END IF






		'''''    ypos  = ypos + 200
		    isect = rs("idsec")

            '     end if

         if   Wtacc = "M" then

              sActcon = "CoprocaccionCons.asp?codobjeto=" & sOCesque & "&Tipobjeto=" & sOTesque
              sActcon = sActcon & "&bversion=" & bVersion & "&baccion=" & rs("CODACCION") & "&nID=" & rs("CODPROCED")
              WACCDEC = rs("CODACCION")

         else

              sActcon = "CoprocaccionaltCons.asp?codobjeto=" & sOCesque & "&Tipobjeto=" & sOTesque & "&bsec=" & rs(0)
              sActcon = sActcon & "&bversion=" & bVersion & "&naccion=" & WACCDEC & "&nID=" & rs("CODPROCED") & "&baccion=" & replace(rs("CODACCION"),"ALT","")
         end if

              sHref = "&nbsp;"


              sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & trim(sOTesque & "") & "'); return false;" & Chr(34)



       select case ucase(rs(3)&"")
         case "INICIO"
           SHTML=SHTML & "<div xonclick='alert(this.clientWidth + "" "" + this.clientHeight + "" "" + this.style.top + "" "" + this.name);'   class='block'  Align=center id='" &rs("CODACCION") &"'"  & " name='" & rs("CODPROCED") &"'"
           SHTML=SHTML &  " style='border:none;background-color: #ffffff;left: "& xpos &"; top: "&ypos&";width:31px;height:30px;padding:0px'>"
          case "FIN"
            SHTML=SHTML & "<div xonclick='alert(this.clientWidth + "" "" + this.clientHeight + "" "" + this.style.top + "" "" + this.name);'  class='block'  Align=center id='" &rs("CODACCION") &"'"  & " name='" & rs("CODPROCED") &"'"
            SHTML=SHTML &  " style='border:none;background-color: #ffffff;left: "& xpos &"; top: "&ypos&";width:31px;height:30px;padding:0px'>"

          CASE "DECISIÓN"
            SHTML=SHTML & "<div " & sFuncion & "  class='block'  Align=center id='" &rs("CODACCION") &"'" & " name='" & rs("CODPROCED") &"'"
            SHTML=SHTML &  " style='cursor:pointer;border:none;background-color: #ffffff;left: "& xpos &"; top: "&ypos&";width:35px;height:34px;padding:0px'>"
        case "PROCESO"
	               SHTML=SHTML & "<div  " & sFuncion & " class='block'  vvalign='middle' vAlign='center'    id='" &rs("CODACCION") &"'" & " name='" & rs("CODPROCED") &"'"
	               'SHTML=SHTML &  " style=""background-repeat:no-repeat;background-image: url('iconos/ICORCDistribucion.jpg');	background-size: cover ;cursor:pointer;left: "& xpos &"; top: "&ypos&";height:50px;font-size: 7pt; font-family: Arial;box-shadow: 4px 0px 5px  #888888 ;"">"
	              ' response.write cstr(rs("archivocur")&"") & "   "
	               if cstr(rs("archivocur")&"") ="" then
	   			    SHTML=SHTML &  " style=""cursor:pointer;left: "& xpos &"; top: "&ypos&";height:50px;font-size: 7pt; font-family: Arial;box-shadow: 4px 0px 5px  #888888 ;"">"
	   			 else
	   			   ' sHTML=SHTML &  "<img   src='iconos/"& rs("archivocur") & "'>"'</td></tr></table>"
	                  if rs("archivocur")<>"DEFAULT" THEN
	                  SHTML=SHTML &  " style=""border-radius:0;border-color:white;background-image: url('iconosCURSOGRAMA/"& rs("archivocur")&"');background-size:100% 100%; ;cursor:pointer;left: "& xpos &"; top: "&ypos&";height:50px;cowlor:"&Session("CORGGFMFT")&";fownt-size:" & Session("CORGGFMFTSZ") & "pt; fwont-family: '"& session("CORGGFMFTFC")&"';"">"
	                  ELSE
	   			   SHTML=SHTML &  " style=""background-COLOR: " & Session("CORGGFM") &" ;cursor:pointer;left: "& xpos &"; top: "&ypos&";height:50px;color:"&Session("CORGGFMFT")&";box-shadow: 4px 0px 5px  #888888 ;"">"
	   			   END IF

	   			     'filter:progid


	                end if

	           case "PROCESOR"
	               SHTML=SHTML & "<div  " & sFuncion & " class='block '  vvalign='middle' vAlign='center'    id='" &rs("CODACCION") &"'" & " name='" & rs("CODPROCED") &"'"
	               'SHTML=SHTML &  " style=""background-repeat:no-repeat;background-image: url('iconos/ICORCDistribucion.jpg');	background-size: cover ;cursor:pointer;left: "& xpos &"; top: "&ypos&";height:50px;font-size: 7pt; font-family: Arial;box-shadow: 4px 0px 5px  #888888 ;"">"
	              ' response.write cstr(rs("archivocur")&"") & "   "
	               if cstr(rs("archivocur")&"") ="" then
	   			    SHTML=SHTML &  " style=""cursor:pointer;left: "& xpos &"; top: "&ypos&";height:50px;font-size: 7pt; font-family: Arial;box-shadow: 4px 0px 5px  #888888 ;"">"
	   			 else
	   			   ' sHTML=SHTML &  "<img   src='iconos/"& rs("archivocur") & "'>"'</td></tr></table>"
	                  if rs("archivocur")<>"DEFAULT" THEN
	                  SHTML=SHTML &  " style=""border-radius:0;border-color:white;background-image: url('iconosCURSOGRAMA/"& rs("archivocur")&"');background-size:100% 100%; ;cursor:pointer;left: "& xpos &"; top: "&ypos&";height:50px;cowlor:"&Session("CORGGFMFT")&";fownt-size:" & Session("CORGGFMFTSZ") & "pt; fwont-family: '"& session("CORGGFMFTFC")&"';"">"
	                  ELSE
	   			   SHTML=SHTML &  " style=""background-COLOR: " & Session("CORGGFM") &" ;cursor:pointer;left: "& xpos &"; top: "&ypos&";height:50px;color:"&Session("CORGGFMFT")&";box-shadow: 4px 0px 5px  #888888 ;"">"
	   			   END IF

	   			     'filter:progid

	   'filter: progid:DXImageTransform.Microsoft.AlphaImageLoader( src='images/img_Header03.png', sizingMethod='scale');:DXImageTransform.Microsoft.AlphaImageLoader(src='images/logo.gif',sizingMethod='scale');-ms-filter:"progid:DXImageTransform.Microsoft.AlphaImageLoader(src='images/logo.gif',sizingMethod='scale')";'
             end if
        end select


      select case ucase(rs(3)&"")
	           case "INICIO"
	                 if cstr(rs("archivocur")&"") = "" OR cstr(rs("archivocur")&"") ="DEFAULT" then
	                 sHTML=SHTML &  "<img   src='iconosCURSOGRAMA/inicio.gif'>"'</td></tr></table>"
	                 else


	                 sHTML=SHTML &  "<img   src='iconosCURSOGRAMA/"& rs("archivocur") & "'>"'</td></tr></table>"
	                 end if
	                 if ("" & rs(13))<> "" then
	  		        SHTMLx=  SHTMLx & "<div class='connector " & rs("CODACCION") & " " & ("" & rs(13)) & "'><img src='JSGRAF/arrow.gif' class='connector-end'></div>"
	                 end if

	          case "FIN"
	                if cstr(rs("archivocur")&"") ="" OR cstr(rs("archivocur")&"") ="DEFAULT" then
	                SHTML=SHTML & "<img   src='iconosCURSOGRAMA/fin.gif'>"
	                else
	                SHTML=SHTML & "<img   src='iconosCURSOGRAMA/"& rs("archivocur") &"'>"
              end if

        CASE "PROCESO"

'''''''''''''replace(rs(1)," ","<br>")

           '  wonka = split(rs(1)," ")
           '  wonkk = ""
           '  wonko = ""


           '    for iii=0 to ubound(wonka)
           '      wonko = wonko & wonkk & wonka(iii)
           '      if wonkk = "&nbsp;" then
           '         wonkk = "<BR>"
           '      else
           '         wonkk = "&nbsp;"
           '      end if
           '    next
           '    if iii=2 then wonko = replace(wonko,"&nbsp;","<BR>")



 ''''''''''style='position:relative;top:0%;'

           'SHTML=SHTML & " <span  xstyle='{position: absolute;top: 30%;left:2%}' aliqgn='center' nmargin=2 >" & wonko & "</span>"
             ' SHTML=SHTML & "<table height=100%  margin=0 cellpadding=0 cellspacing=0 style='text-align:center;font-size:7pt;font-family: Arial;'><tr  ><td valign=middle>" & wonko & "</td></tr></table>"
               SHTML=SHTML & "<table    height=100%  margin=0 cellpadding=0 cellspacing=0 style='text-align:center;'><tr  ><td   valign=middle><font style='font-size:"&Session("CORGGFMFTSZ") &"px' color=" &Session("CORGGFMFT")&" face=" & session("CORGGFMFTFC") & ">" & wonko & "</font></td></tr></table>"

              if ("" & rs(13))<> "" then
                    IF Wtacc = "M" then
		       SHTMLx=  SHTMLx & "<div class='connector " & rs("CODACCION") & " " & ("" & rs(13)) & "'><img src='JSGRAF/arrow.gif' class='connector-end'></div>"
                    else
''b.id as SECSIid , c.id as SECNOid
                       WtaccN = "A"
                       if UCASE(left(("" & rs(13)),3) & "") <> "ALT" THEN
                          WtaccN = "M"
                        END IF

                        if WtaccN = "A" then
		           SHTMLx=  SHTMLx & "<div  class='connector " & rs("CODACCION") & " " & ("" & rs(13)) & "'><img src='JSGRAF/arrow.gif' class='connector-end'></div>"
                        else
                           if rs("SECSIid") > rs(0) then
		              SHTMLx=  SHTMLx & "<div class='connector " & rs("CODACCION") & " " & ("" & rs(13)) & "  right_start up_end'><img src='JSGRAF/arrow.gif' class='connector-end'></div>"
                           else
		              SHTMLx=  SHTMLx & "<div class='connector " & rs("CODACCION") & " " & ("" & rs(13)) & "  down_start up_end'><img src='JSGRAF/arrow.gif' class='connector-end'></div>"
                           end if

                        end if

                    end if
              end if
 CASE "PROCESOR"


              SHTML=SHTML & "<table    height=100%  margin=0 cellpadding=0 cellspacing=0 style='text-align:center;'><tr  ><td   valign=middle><font style='font-size:"&Session("CORGGFMFTSZ") &"px' color=" &Session("CORGGFMFT")&" face=" & session("CORGGFMFTFC") & ">" & wonko & "</font></td></tr></table>"


              if ("" & rs(13))<> "" then
                    IF Wtacc = "M" and rs(13) <> "NO" then
		       SHTMLx=  SHTMLx & "<div class='connector " & rs("CODACCION") & " " & ("" & rs(13)) & "'><img src='JSGRAF/arrow.gif' class='connector-end'></div>"
                    else
''b.id as SECSIid , c.id as SECNOid
                       WtaccN = "A"
                       if UCASE(left(("" & rs(13)),3) & "") <> "ALT" and rs(13) <> "NO" THEN
                          WtaccN = "M"
                        END IF

                        if WtaccN = "A" and rs(13) <> "NO" then
		           SHTMLx=  SHTMLx & "<div  class='connector " & rs("CODACCION") & " " & ("" & rs(13)) & "'><img src='JSGRAF/arrow.gif' class='connector-end'></div>"
                        else
                           if rs("SECSIid") > rs(0) and rs(13) <> "NO" then
		              SHTMLx=  SHTMLx & "<div class='connector " & rs("CODACCION") & " " & ("" & rs(13)) & "  right_start up_end'><img src='JSGRAF/arrow.gif' class='connector-end'></div>"
                           else
                           if  rs(13) <> "NO" then
		              SHTMLx=  SHTMLx & "<div class='connector " & rs("CODACCION") & " " & ("" & rs(13)) & "  down_start up_end'><img src='JSGRAF/arrow.gif' class='connector-end'></div>"
		              end if
                           end if

                        end if

                    end if
              end if
             'ALINKC(IALINKC,0)=rs("CODACCION") & ""
            ' ALINKC(IALINKC,1)=rs(13) & ""
             IALINKC=IALINKC +1
             sidcon= sidcon & "<span  name='" &  rs("CODACCION") &"' class=xcon style='kposition:absolute;background-Color:#01A9DB;' id='" & IALINKC  &"O'><font color=#ffffff>"  &  IALINKC & "</font></span>"
             sidcon= sidcon & "<span  name='" &  rs("codaccioncn") &"' class=xcon style='kposition:absolute;;background-Color:#FF4000;' id='" & IALINKC  & "D'><font color=#ffffff>"  &  IALINKC & "</font></span>"

CASE "DECISIÓN"

''''''''''''''''''replace(rs(1)," ","&nbsp;")

              'shTML=SHTML & "<img  src='JSGRAF/decision.gif'>"
               if cstr(rs("archivocur")&"") ="DEFAULT" OR cstr(rs("archivocur")&"") ="" then
                SHTML=SHTML & "<img   src='iconosCURSOGRAMA/decision.gif'>"
              else
                SHTML=SHTML & "<img   src='iconosCURSOGRAMA/"& rs("archivocur") &"'>"
              end if
              if rs("SECSIid") < rs("id") then
                SHTMLx=  SHTMLx & "<div class='connector " & rs("CODACCION") & " " & ("" & rs(13))  & " left_start up_end'><img src='JSGRAF/arrow.gif' class='connector-end'><label class='middle-label'  style='background-color:"& Session("CORGGFN") & ";' ><font style='font-size:"&Session("CORGGFNFTSZ") &"px' color=" &Session("CORGGFNFT")&" face=" & session("CORGGFNFTFC") & ">SI</font></label></div>"
              else
 	        SHTMLx=  SHTMLx & "<div class='connector " & rs("CODACCION") & " " & ("" & rs(13)) & "'><img src='JSGRAF/arrow.gif' class='connector-end'><label class='middle-label' style='background-color:"& Session("CORGGFN") & ";'  ><font style='font-size:"&Session("CORGGFNFTSZ") &"px' color=" &Session("CORGGFNFT")&" face=" & session("CORGGFNFTFC") & ">SI</font></label></div>"
              end if


	      if cstr(rs("CODACCIONno")&"") <> "" then
		 if  rs("SECnoid") < rs("id") then
	             SHTMLx=  SHTMLx & "<div class='connector " & rs("CODACCION") & " " & ("" & rs(14)) & " down_start down_end'><img src='JSGRAF/arrow.gif' class='connector-end'><label class='middle-label' style='background-color:"& Session("CORGGFN") & ";'><font style='font-size:"&Session("CORGGFNFTSZ") &"px' color=" &Session("CORGGFNFT")&" face=" & session("CORGGFNFTFC") & ">NO</font></label><label  id='xxr' style='z-index:9999;' class='connector-start'><font style='font-size: 7pt; font-family: Arial;'> " & wonko & "?</font></label></div>"

	         else
		    SHTMLx=  SHTMLx & "<div class='connector " & rs("CODACCION") & " " & ("" & rs(14)) & " down_start right_end'><img src='JSGRAF/arrow.gif' class='connector-end'><label class='middle-label' style='background-color:"& Session("CORGGFN") & ";'><font style='font-size:"&Session("CORGGFNFTSZ") &"px' color=" &Session("CORGGFNFT")&" face=" & session("CORGGFNFTFC") & ">NO</font></label><label id='xxr' style='z-index:9999;background-color:"& Session("CORGGFN") & ";' class='connector-start'> <font style='font-size:"&Session("CORGGFNFTSZ") &"px' color=" &Session("CORGGFNFT")&" face=" & session("CORGGFNFTFC") & ">" & wonko & "?</font></label></div>"

	         end if
	         IF RS("SECSIid") <> "" AND RS("SECNOid") & "" <> "" THEN
		    IF RS("SECSIid") &"" <> RS("SECNOid") & "" THEN
		       XS1=RS("SECSI")
		       XS2=RS("SECSI")
		       XPOSDES=XPOS
		    END IF
		 ELSE
		   'XPOSDES=XPOS
	         END IF
              end if

       end select


      SHTML=SHTML &  " </div>"


      ASEC(I,0)=xcols
	  ASEC(I,1)=xpos
	  ASEC(I,2)=ypos
      ASEC(I,3)=RS("IDSEC")
      i=i+1
        if clng("0" & rs("posx" )) <= 0 then
			    sql="update OBJPROCURSODET  set posx='" & xpos & "',posy='" & ypos & "' where id=" & rs(0)
			    set re=cn.execute(sql)
	    end if
	'''    xpos = xpos + 150

          xposant = XPOS

          XPOS = XPOSPREV





      rs.movenext
      loop
             IF IIPRC > 0 THEN
                APRCEV(IIPRC,3) = xposant
             END IF





spos="<script>function poselm(){" & CHR(13)
sql="select CODACCION,posx,posy from OBJPROCURSODET where codobjeto = " & sOCesque & "  "
set re=cn.execute(sql)

do while not re.eof

spos=SPOS & " document.getElementById ('" & re(0) &"').style.left='"  & re(1) & "px';" &chr(13)

spos=SPOS & " document.getElementById ('" &  re(0) &"').style.top='"  & re(2) & "px';" &chr(13)
re.movenext
loop
spos=SPOS & " ver4(document);initPageObjects();repaintcnn();document.getElementById('mainCanvas').style.visibility='visible';  }</script>"

RESPONSE.WRITE SHTML &SHTMLx   & sidcon




      %>



     </div>


     <%
    RESPONSE.WRITE  spos
     %>






<%
       if muestra="S" then
         'response.write "poselm();"
       end if

%>

     <script type="text/javascript">

function ver4(t)
     {

     var dato='';
     var prca='N';
     var prcn='N';
     var prcw=25;
     var prct=0;
     var WmaxH=80;
     var WmaxW=500;
     var zz= document.getElementById('mainCanvas');

     var ele=getElementsByClassName(zz,'block');



     for(i=0;i< ele.length;i++)
     {
      posx=document.getElementById (ele[i].id).style.left;
      posy=document.getElementById (ele[i].id).style.top;
      posH=parseInt(document.getElementById (ele[i].id).clientHeight);

      Wposx=parseInt(document.getElementById (ele[i].id).style.left);




      if (Wposx > WmaxW) {WmaxW = Wposx;}

      xxx1 = 0;
      xxx2 = 0;

      xxx1 = Math.round(posH/2);
      xxx2 = parseInt(posH/2);

     if ( xxx1 != xxx2 ) {
          posH=posH+1;
     //     document.getElementById (ele[i].id).height= clientHeight + 'px';
     //     alert(ele[i].id + '???' + document.getElementById (ele[i].id).clientHeight + '???' + document.getElementById (ele[i].id).style.height);
       }

      if (posH > WmaxH) {WmaxH = posH;}
      dato=dato + ele[i].id + '###'  +   posx + '###' + posy  + '###' + posH + '///'


   //   alert(ele[i].id + ' ' + document.getElementById (ele[i].id).style.left + ' ' + document.getElementById (ele[i].id).offsetLeft);

      Wposc=parseInt(document.getElementById (ele[i].id).style.left);
      Wposw=parseInt(document.getElementById (ele[i].id).clientHeight);

      prcn=document.getElementById (ele[i].id).attributes['name'].value;
  //   alert(document.getElementById (ele[i].id).attributes['name'].value);
      if (prca!=prcn) {

         if (prca!='N') {
             document.getElementById('main' + prca).style.width  = (Wposc - prcw) + 'px';
             prct=prct + (Wposc - prcw);
             prcw=Wposc;

            }
          prca=prcn;

        }


     }


   //  alert('1');
   //  alert((Wposc - prcw + (Wposw * 2)) + 'px');
   //  alert(prca);
   //  alert(document.getElementById('main' + prca).id);
      document.getElementById('main' + prca).style.width  = (Wposc - prcw + (Wposw * 2)) + 'px';
   //  alert('2');
             prct=prct + (Wposc - prcw + (Wposw * 2));
   //  alert('3');


  //     alert(WmaxH + '???' + dato);



     for(i=0;i< ele.length;i++)
     {
      posH=document.getElementById (ele[i].id).clientHeight;
      Pheight = parseInt(posH);
      posHK = (WmaxH - Pheight) / 2;
      document.getElementById (ele[i].id).name= posHK;
      pagrK = 0;
      xxx1 = Math.round(Pheight/2);
      xxx2 = parseInt(Pheight/2);



      if ( Pheight < 42 ) {
          pagrK=1;

       }

      posHK = Math.round((WmaxH - Pheight) / 2) + pagrK ;
      document.getElementById (ele[i].id).name= document.getElementById (ele[i].id).name + '-' + posHK;
  //     alert(posHK + '???');
      Pheight = parseInt(document.getElementById (ele[i].id).style.top);
      document.getElementById (ele[i].id).style.top= (Pheight + posHK) + 'px';



     }

//     alert('2');

    //  document.getElementById('mainCanvas').style.width = (WmaxW + 60) + 'px';

   //   document.getElementById('mainTabla').style.width  = (WmaxW + 250) + 'px';


 document.getElementById('mainCanvas').style.height=window.screen.availHeight-150 + 'px';
   //   document.getElementById('mainCanvas').style.width = prct + 'px';
//alert(document.getElementById('mainCanvas').style.width);
     document.getElementById('mainTabla').style.width  = prct + 'px';


     }







     </script>
<script>
//CursorDivScroll( 'mainCanvas');
function getElementsByClassName(node, classname) {
			     var a = [];
			     var re = new RegExp('(^| )'+classname+'( |$)');
			     var els = node.getElementsByTagName("*");
			     for(var i=0,j=els.length; i<j; i++)
			         if(re.test(els[i].className))a.push(els[i]);
			     return a;
}
function drawCanvas(texto,imgid) {

	 		  		   var canvas = document.createElement('canvas');

	 		  		    <%
					   if session("EMPMUN") = "S" and objun = "1" then
					   	     response.write "canvas.width='35';"
					   else
					   		 response.write "canvas.width='15';"
					   end if
	 		  		   %>

	 		  		   canvas.height='295';
	                     var img = document.getElementById(imgid);
                      // alert(canvas.getContext);
	 		  		   var context = canvas.getContext('2d');

	 		  		 // Fill Stroke Text
	                    var x = canvas.width / 2;
	       				var y = canvas.height / 2;
	 		  		   context.lineWidth=1;

	 		  		   context.strokeStyle="rgba(35, 70, 237, .8)";
	                   //  context.translate(70 / 2, 40 / 2);
	 		  		   context.font="10pt arial";
	 		  		    context.translate(x ,y);
	                    context.rotate(270*Math.PI/180);

	                    context.textAlign = 'center';
	 		  		   <%
					   if session("EMPMUN") = "S" and objun = "1" then
					   		response.write "var t1 = texto.split('-');"
					   		response.write "context.fillText(t1[0], 0, 0);"
					   		response.write "context.fillText(t1[1], 0, 12);"
					   else
					   		response.write "context.fillText(texto, 0, 4);"
					   end if
	 		  		   %>

	 		  		//   context.strokeText("Fill, Stroke Text, 28pt, sans-serif", 0, 0);


	          img.src=canvas.toDataURL("")
	 		  		 }

function getInternetExplorerVersion()
// Returns the version of Internet Explorer or a -1
// (indicating the use of another browser).
{
  var rv = -1; // Return value assumes failure.
  if (navigator.appName == 'Microsoft Internet Explorer')
  {
    var ua = navigator.userAgent;
    var re  = new RegExp("MSIE ([0-9]{1,}[\.0-9]{0,})");
    if (re.exec(ua) != null)
      rv = parseFloat( RegExp.$1 );
  }
  return rv;
}

if (getInternetExplorerVersion() > 8 || getInternetExplorerVersion()==-1)
{

var ele=getElementsByClassName( document.body,'secox');

				      for(i=0;i< ele.length;i++)
				      {
				       title= document.getElementById (ele[i].id).title;

				       drawCanvas(title,ele[i].id);
				       document.getElementById (ele[i].id).style.display='';

}
				        var ele=getElementsByClassName( document.body,'spanxx');

					   				      for(i=0;i< ele.length;i++)
					   				      {
					   				     // alert(document.getElementById (ele[i].id).outerHTML);
					   				    document.getElementById (ele[i].id).outerHTML ='';

}

}
</script>


<script>
function xposConectores()
{

var ele = getElementsByClassName(document,'xcon');

  for(i=0;i< ele.length;i++)
  {

  var div =document.getElementById(ele[i].getAttribute('name') + 'C');
  if (div ==null)
  {

  var div = document.createElement("div");
  var doc=document.getElementById('mainCanvas');
  doc.appendChild(div);
  div.setAttribute("id", ele[i].getAttribute('name') + 'C');
 // div.style.border="thick solid #0000FF";

  div.style.position='absolute';

  }
  //+ Math.round( parseInt(document.getElementById(ele[i].getAttribute('name')).clientWidth  /2))-5 ;
	x=Math.round(parseInt(document.getElementById(ele[i].getAttribute('name')).style.left)) ;
	y=Math.round(parseInt(document.getElementById(ele[i].getAttribute('name')).style.top)) + Math.round( parseInt(document.getElementById(ele[i].getAttribute('name')).clientHeight  + 10 ));
	//document.getElementById (ele[i].id).style.left=x + 'px';
	//document.getElementById (ele[i].id).style.top=y + 'px';

	div.style.left=x +3 + 'px';
    div.style.top=y  + 'px';
    div.innerHTML=div.innerHTML   + ele[i].outerHTML ;
    if ((div.childNodes.length %5)==4)
    {
     div.innerHTML=div.innerHTML + '<br><br>'
    }

    ele[i].outerHTML='';


	}

 }
</script>
       </body>

       </html>




<%


       end if

''response.write sHTML

%>