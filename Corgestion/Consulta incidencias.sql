select a.id,b.descripcion,c.descripcion,f.descripcion,g.descripcion,a.fecha_ingreso,a.respon_gestion,a.respon_tec,h.nombre,j.nombre,a.CREACION,a.titulo,i.descripcion
from incidentes a 
	inner join clientes b on (a.id_cliente=b.id) 
	inner join prioridades c on (a.prioridad=c.id) 
	inner join estados f on (a.estado_desa=f.id) 
	inner join estados g on (a.estado_func=g.id) 
	inner join cor_responsables h on (a.respon_gestion = h.id) 
	inner join aplicativos i on (a.ID_APLIC = i.ID)  
	left join cor_responsables j on (a.respon_tec = j.id) 
	where a.ver ='1' 