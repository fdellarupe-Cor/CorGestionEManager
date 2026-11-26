<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<html>


<FORM  name="Subconsulta"    action="http://SERVER-41/CORESV3/COREManager/ListadoConsultas.asp?Menu=PORESTADO&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=ESTADO"  method=post >
   <input  type='hidden'  size=60%  value='PORESTADO' name="Tcons" >
   <input  type='hidden'  size=60%  value='1' name="buscar" >
   <input  type=SUBMIT  value='Buscar'>
</FORM>
  <script>document.Subconsulta.submit();</script>
</html>