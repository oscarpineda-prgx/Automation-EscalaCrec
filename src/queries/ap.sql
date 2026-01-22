SELECT 
    YEAR(InvDt)  AS anio,
    MONTH(InvDt) AS mes,
    SUM(ISNULL(GrsInvAmt, 0)) AS GrsInvAmt
FROM SORIANA_PROJECTS.dbo.F_AP(?, ?, ?)
WHERE BSAK_BSIK_XREF3 LIKE ?
GROUP BY 
    YEAR(InvDt),
    MONTH(InvDt)
ORDER BY 
    anio,
    mes;