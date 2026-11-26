const fs=require('fs');
const rd=p=>fs.readFileSync(p,'latin1'), wr=(p,s)=>fs.writeFileSync(p,s,'latin1');
const cnt=(s,x)=>s.split(x).length-1;
const rep=(s,x,y,n,tag)=>{const c=cnt(s,x);if(c!==n)throw new Error(tag+': esperaba '+n+' y hay '+c+' -> '+x.slice(0,80));return s.split(x).join(y)};

// ---------- funciones.inc: funciones nuevas
let f=rd('IncFile/funciones.inc');
if(f.includes('CgComboResponsable'))throw new Error('funciones.inc ya tiene CgComboResponsable');
const addF=[
"",
"''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''",
"' FILTRO POR RESPONSABLE EN LISTADOS (incidentes / proyectos)",
"' Por defecto: registros del usuario de la sesion. Valor \"*\" = todos los registros.",
"'  CgRespSeleccionado() : valor elegido (form o querystring); si no hay, el usuario de la sesion",
"'  CgRespParam(r)       : valor para @Respon (\"\" = sin filtro)",
"'  CgComboResponsable   : <select name=resp> con TODOS, MIS REGISTROS y la lista de responsables",
"''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''",
"Function CgRespSeleccionado()",
"\tDim r",
"\tr = Request.Form(\"resp\")",
"\tIf r & \"\" = \"\" Then r = Request.QueryString(\"resp\")",
"\tIf r & \"\" = \"\" Then r = Session(\"USUARIO\") & \"\"",
"\tCgRespSeleccionado = r",
"End Function",
"",
"Function CgRespParam(r)",
"\tIf r & \"\" = \"*\" Then CgRespParam = \"\" Else CgRespParam = r & \"\"",
"End Function",
"",
"Function CgComboResponsable(cn, seleccionado, formId)",
"\tDim h, rsR, miUsr, v",
"\tmiUsr = Trim(Session(\"USUARIO\") & \"\")",
"\th = \"<select class='ui-select' onchange=\"\"document.getElementById('\" & formId & \"').submit();\"\" id='resp' name='resp'>\"",
"\tIf seleccionado & \"\" = \"*\" Then",
"\t\th = h & \"<option value='*' selected>RESPONSABLE: TODOS</option>\"",
"\tElse",
"\t\th = h & \"<option value='*'>RESPONSABLE: TODOS</option>\"",
"\tEnd If",
"\tIf miUsr <> \"\" And miUsr <> \"0\" Then",
"\t\th = h & \"<option value='\" & Server.HTMLEncode(miUsr) & \"'\"",
"\t\tIf LCase(seleccionado & \"\") = LCase(miUsr) Then h = h & \" selected\"",
"\t\th = h & \">MIS REGISTROS (\" & Server.HTMLEncode(miUsr) & \")</option>\"",
"\tEnd If",
"\tSet rsR = cn.Execute(\"sp_cg_Responsables @Modo='combo_todos'\")",
"\tDo While Not rsR.EOF",
"\t\tv = rsR(0) & \"\"",
"\t\tIf LCase(v) <> LCase(miUsr) Then",
"\t\t\th = h & \"<option value='\" & Server.HTMLEncode(v) & \"'\"",
"\t\t\tIf LCase(seleccionado & \"\") = LCase(v) Then h = h & \" selected\"",
"\t\t\th = h & \">\" & Server.HTMLEncode(rsR(1) & \"\") & \"</option>\"",
"\t\tEnd If",
"\t\trsR.MoveNext",
"\tLoop",
"\trsR.Close",
"\tSet rsR = Nothing",
"\tCgComboResponsable = h & \"</select>\"",
"End Function",
""].join("\r\n");
if(!/%>\s*$/.test(f))throw new Error('funciones.inc no cierra con %>');
f=f.replace(/%>\s*$/,addF+"\r\n%>");

// ---------- datosDetalles.inc
let d=rd('IncFile/datosDetalles.inc');
d=rep(d,'Function GetProyectosListado(cn, IdCliente, IdAplicativo, TextoBusqueda, EstadoFiltro, FechaIngreso, Orden, RegistroInicial, CantidadMostrar)','Function GetProyectosListado(cn, IdCliente, IdAplicativo, TextoBusqueda, EstadoFiltro, FechaIngreso, Orden, RegistroInicial, CantidadMostrar, Respon)',1,'dd1');
d=rep(d,'    cmd.Parameters.Append cmd.CreateParameter("@CantidadMostrar", adInteger, adParamInput, , Clng(CantidadMostrar))','    cmd.Parameters.Append cmd.CreateParameter("@CantidadMostrar", adInteger, adParamInput, , Clng(CantidadMostrar))\r\n    \' [CG] filtro por responsable (usuario de red); "" = todos\r\n    If Trim(Respon & "") = "" Then\r\n        cmd.Parameters.Append cmd.CreateParameter("@Respon", adVarChar, adParamInput, 100, Null)\r\n    Else\r\n        cmd.Parameters.Append cmd.CreateParameter("@Respon", adVarChar, adParamInput, 100, CStr(Respon))\r\n    End If',1,'dd2');

// ---------- incidentes.asp
let a=rd('incidentes.asp');
a=rep(a,'Function ObtenerListadoIncidentes(cn, idCli, idApli, titulo, orden)','Function ObtenerListadoIncidentes(cn, idCli, idApli, titulo, orden, respon)',1,'i1');
a=rep(a,'    cmd.Parameters.Append cmd.CreateParameter("@Orden", 200, 1, 20, CStr(orden & ""))\r\n\r\n    Set ObtenerListadoIncidentes','    cmd.Parameters.Append cmd.CreateParameter("@Orden", 200, 1, 20, CStr(orden & ""))\r\n    \' [CG] filtro por responsable (usuario de red); "" = todos\r\n    If Trim(respon & "") = "" Then\r\n        cmd.Parameters.Append cmd.CreateParameter("@Respon", 200, 1, 100, Null)\r\n    Else\r\n        cmd.Parameters.Append cmd.CreateParameter("@Respon", 200, 1, 100, CStr(respon))\r\n    End If\r\n\r\n    Set ObtenerListadoIncidentes',1,'i2');
a=rep(a,'    Set rsListado = ObtenerListadoIncidentes(Cnn, cliente, aplicativo, tit, ord)','    Dim cgResp\r\n    cgResp = CgRespSeleccionado() \' [CG] por defecto: registros del usuario de la sesion; "*" = todos\r\n    Set rsListado = ObtenerListadoIncidentes(Cnn, cliente, aplicativo, tit, ord, CgRespParam(cgResp))',1,'i3');
a=rep(a,'    rsApli.Close\r\n    sPage = sPage & "                        </select>" & vbCrLf\r\n    sPage = sPage & "                    </form>" & vbCrLf','    rsApli.Close\r\n    sPage = sPage & "                        </select>" & vbCrLf\r\n    \' [CG] Combo Responsable (por defecto: usuario de la sesion)\r\n    sPage = sPage & "                        " & CgComboResponsable(Cnn, cgResp, "xForm") & vbCrLf\r\n    sPage = sPage & "                    </form>" & vbCrLf',1,'i4');
let L=a.split('\r\n'),k=0;
const pat='&cli=" & cliente & "&xmenu="';
for(let i=0;i<L.length;i++){ if(L[i].includes('incidentes.asp?')&&!L[i].includes('opc=consi')&&L[i].includes(pat)){L[i]=L[i].split(pat).join('&cli=" & cliente & "&resp=" & Server.URLEncode(cgResp) & "&xmenu="');k++} }
if(k!==5)throw new Error('incidentes: URLs actualizadas '+k+' (esperaba 5)');
a=L.join('\r\n');

// ---------- proyectos.asp
let p=rd('proyectos.asp');
const eol=p.includes('\r\n')?'\r\n':'\n';
p=rep(p,'usuario = userlogon()','usuario = userlogon()'+eol+"' [CG] Filtro por responsable: por defecto los registros del usuario de la sesion; \"*\" = todos. En consulta (consi) no se filtra."+eol+'If IsConsulta Then cgResp = "*" Else cgResp = CgRespSeleccionado()',1,'p1');
p=rep(p,'GetProyectosListado(scon1 , cliente, aplicativos, tit, estadoFiltro, fechaFiltro, ord, CantRegMover, CantRegMostrar)','GetProyectosListado(scon1 , cliente, aplicativos, tit, estadoFiltro, fechaFiltro, ord, CantRegMover, CantRegMostrar, CgRespParam(cgResp))',1,'p2');
p=rep(p,'&xtmenu=" & xtmenu & "&apli=" & aplicativos & "\'>"','&xtmenu=" & xtmenu & "&apli=" & aplicativos & "&resp=" & Server.URLEncode(cgResp) & "\'>"',1,'p3');
p=rep(p,'&xtmenu=" & xtmenu & "&cli=" & cliente & "\'>"','&xtmenu=" & xtmenu & "&cli=" & cliente & "&resp=" & Server.URLEncode(cgResp) & "\'>"',1,'p4');
const marker='    html = html & "</select></form></td>"'+eol+eol+'    \' Bot';
const nuevo='    html = html & "</select></form></td>"'+eol+eol+'    \' [CG] Form Filtro Responsable (por defecto: usuario de la sesion)'+eol+'    html = html & "<td><form id=\'xFormResp\' method=\'post\' action=\'proyectos.asp?ord=" & ord & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cli=" & cliente & "&apli=" & aplicativos & "\'>" & CgComboResponsable(Cnn, cgResp, "xFormResp") & "</form></td>"'+eol+eol+'    \' Bot';
p=rep(p,marker,nuevo,1,'p5');
p=rep(p,'UrlBaseParams = "&cli=" & cliente & "&apli=" & aplicativos & "&xmenu="','UrlBaseParams = "&cli=" & cliente & "&apli=" & aplicativos & "&resp=" & Server.URLEncode(cgResp) & "&xmenu="',1,'p6');
p=rep(p,'UrlPagiBase = "proyectos.asp?ord=" & ord & "&cli=" & cliente & "&apli=" & aplicativos & "&xmenu="','UrlPagiBase = "proyectos.asp?ord=" & ord & "&cli=" & cliente & "&apli=" & aplicativos & "&resp=" & Server.URLEncode(cgResp) & "&xmenu="',1,'p7');

// todo validado: escribir
wr('IncFile/funciones.inc',f);wr('IncFile/datosDetalles.inc',d);wr('incidentes.asp',a);wr('proyectos.asp',p);
console.log('ok');
