WITH sumatoria (doc,codigo,suma) as 
(
select c5_crfndoc,c6_ccodigo,sum(c6_ncantid) from al0001movd a
inner join al0001movc b on a.c6_calma=b.c5_calma and a.c6_ctd=b.c5_ctd and a.c6_cnumdoc=b.c5_cnumdoc
where b.C5_CSITUA<>'A' AND a.C6_DFECDOC<=@fechaguia
group by c5_crfndoc,c6_ccodigo
),
 facturas(fnumser,fdoc,ftdoc,fcodigo,fsuma) AS
 (
 SELECT f6_cnumser,f6_cnumdoc,f6_ctd,f6_ccodigo,sum(f6_ncantid) FROM FT0001ACUd
 WHERE F6_CESTADO<>'A' 
 GROUP BY f6_cnumser,f6_cnumdoc,f6_ccodigo,f6_ctd
 UNION ALL
 SELECT f6_cnumser,f6_cnumdoc,f6_ctd,f6_ccodigo,sum(f6_ncantid) FROM FT0001facd
 WHERE F6_CESTADO<>'A' 
 GROUP BY f6_cnumser,f6_cnumdoc,f6_ccodigo,f6_ctd
 
  )

select f.f6_ccodage TIENDA,f.f6_ctd [TIPO_DOC],LTRIM(RTRIM(f.f6_cnumser))+LTRIM(RTRIM(f.f6_cnumdoc)) DOC_DIFERIDO,F6_DFECDOC FECHA_EMI,f5_ccodcli CLIENTE,f.f5_cnombre NOMBRE,
F.f6_ccodigo COD_PROD,f.f6_cdescri PRODUCTO,F.F5_CCODMON MONEDA,F.F6_NPRECIO [PRECIO_UNITARIO],CASE WHEN F.F5_CCODMON ='US' THEN F.F6_NIMPUS ELSE F.F6_NIMPMN END TOTAL,
fac.fsuma CANT_FACTURADA,coalesce(fac.fsuma-s.suma,fac.fsuma) as  SALDO_TOTAL_X_ATENDER,
case when g.c6_ncantid is null then 0 
WHEN G.C6_DFECDOC>@fechaguia AND coalesce(f.f6_ncantid-s.suma,0)>=0  THEN 0
else g.c6_ncantid end CANT_ATENDIDA,
case when g.c6_ctd is null then CASE WHEN f.F6_NSALDAR =0 THEN 'ATENDIDO X BD (NC)' ELse 'X ATENDER'
END WHEN G.C6_DFECDOC>@fechaguia AND coalesce(f.f6_ncantid-s.suma,0)>=0  THEN 'X ATENDER' else g.c6_ctd end [TIPO_DOC_A],
case when g.c6_ctd is null then CASE WHEN f.F6_NSALDAR =0 THEN 'ATENDIDO X BD (NC)' else 'X ATENDER' 
end WHEN G.C6_DFECDOC>@fechaguia AND coalesce(f.f6_ncantid-s.suma,0)>=0  THEN 'X ATENDER' else g.c6_cnumdoc end DOCUMENTO_ATIENDE,
case when g.c6_ctd is null then CASE WHEN f.F6_NSALDAR =0 THEN '1900-01-01' else '1900-01-01'
end WHEN G.C6_DFECDOC>@fechaguia AND coalesce(f.f6_ncantid-s.suma,0)>=0  THEN '1900-01-01' else g.c6_DFECDOC end [FECHA_GUIA],
case when g.c6_clocali is null then CASE WHEN f.F6_NSALDAR =0 THEN 'ATENDIDO X BD (NC)' else 'X ATENDER'
end WHEN G.C6_DFECDOC>@fechaguia AND coalesce(f.f6_ncantid-s.suma,0)>=0  THEN 'X ATENDER' else g.c6_clocali end TIENDA_ATIENDE,
case when g.c6_calma is null then CASE WHEN f.F6_NSALDAR =0 THEN 'ATENDIDO X BD (NC)' else 'X ATENDER'
end WHEN G.C6_DFECDOC>@fechaguia AND coalesce(f.f6_ncantid-s.suma,0)>=0  THEN 'X ATENDER' else g.c6_calma end ALMACEN_ATIENDE
from 
(
select fd.*,FC.F5_CTF,FC.F5_CESTADO,fc.f5_ccodcli,fc.f5_cnombre,F5_CCODMON
from
(select f6_ccodage,f6_ctd,f6_cnumser,f6_cnumdoc,f6_dfecdoc,f6_ccodigo,f6_cdescri,f6_nprecio,f6_nimpus,f6_nimpmn,sum(f6_ncantid) f6_ncantid,f6_nsaldar from FT0001ACUd 
GROUP BY f6_ccodage,f6_ctd,f6_cnumser,f6_cnumdoc,f6_dfecdoc,f6_ccodigo,f6_cdescri,f6_nprecio,f6_nimpus,f6_nimpmn,f6_ncantid,f6_nsaldar
union all 
select f6_ccodage,f6_ctd,f6_cnumser,f6_cnumdoc,f6_dfecdoc,f6_ccodigo,f6_cdescri,f6_nprecio,f6_nimpus,f6_nimpmn,sum(f6_ncantid) f6_ncantid,f6_nsaldar from ft0001facd
GROUP BY f6_ccodage,f6_ctd,f6_cnumser,f6_cnumdoc,f6_dfecdoc,f6_ccodigo,f6_cdescri,f6_nprecio,f6_nimpus,f6_nimpmn,f6_ncantid,f6_nsaldar
) fd left join
(select * from FT0001acuc union all select * from ft0001facc) fc 
on f5_ccodage=f6_ccodage and f5_ctd=f6_ctd and f5_cnumser=f6_cnumser and f5_cnumdoc=f6_cnumdoc
where f5_ctf='05' 
and F5_CESTADO<>'A'
) f
left join
(
select c5_ccodmov,C5_CSITUA,c5_crftdoc,c5_crfndoc,c6_ccodigo,sum(c6_ncantid) c6_ncantid,c6_dfecdoc,c6_ctd,c6_cnumdoc,c6_clocali,c6_calma from al0001movd d left join al0001movc c on c6_calma=c5_calma and c6_ctd=c5_ctd and c6_cnumdoc=c5_cnumdoc
where c5_ccodmov in ('08','81') and C5_CSITUA<>'A'
GROUP BY c5_ccodmov,C5_CSITUA,c5_crftdoc,c5_crfndoc,c6_ccodigo,c6_dfecdoc,c6_ctd,c6_cnumdoc,c6_clocali,c6_calma
) g 
on f6_ctd=c5_crftdoc and f6_cnumser+f6_cnumdoc=c5_crfndoc and F6_CCODIGO=C6_CCODIGO
left join sumatoria s on g.c6_ccodigo=s.codigo and g.c5_crfndoc=s.doc
LEFT JOIN facturas fac ON f6_ctd=fac.ftdoc and f6_cnumser=fac.fnumser AND f6_cnumdoc=fac.fdoc and C6_CCODIGO=fac.Fcodigo
where not F6_CCODAGE is null