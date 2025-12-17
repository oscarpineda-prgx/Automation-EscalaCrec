SELECT acuerdo, clase_acuerdo, descripción_del_acuerdo, denominación_externa, no_proveedor, fecha_modificación_cancelación
FROM SORIANA_PROJECTS.dbo.F_CONVENIOS_SAP(?)
WHERE
    UPPER([descripción_del_acuerdo]) LIKE '%ESCALA%';