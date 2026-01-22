SELECT 
    VndNbr  AS [N° Proveedor],
    VndName AS [Nombre Proveedor]
FROM SORIANA_PROJECTS.dbo.P_VendorMaster
WHERE VndNbr = ?
ORDER BY VndNbr;
