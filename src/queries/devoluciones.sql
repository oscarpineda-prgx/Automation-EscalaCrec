SELECT
    YEAR(COALESCE(rcvdt, podt))  AS anio,
    MONTH(COALESCE(rcvdt, podt)) AS mes,
    SUM(ISNULL([compra neta mas impuestos], 0)) AS devoluciones_mas_impuestos
FROM dbo.F_DEVOLUCIONES(?, ?, ?)
WHERE ISNULL(po_org, 'SORIANA') <> 'CITY'
GROUP BY
    YEAR(COALESCE(rcvdt, podt)),
    MONTH(COALESCE(rcvdt, podt))
ORDER BY
    anio, mes;