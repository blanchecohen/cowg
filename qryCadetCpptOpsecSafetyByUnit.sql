SELECT X.*
FROM (SELECT N.Total, N.CPPTcnt AS CPPT, N.OPSECcnt AS OPSEC, N.SafetyCnt AS Safety, W.UnitName, W.Unit
FROM (WingUnits AS W inner join SubUnitsWithCadets as C on W.Unit=C.unit) LEFT JOIN qryCadetCpptOPSECCountsSUB AS N ON N.Unit=W.Unit
Where W.Unit<>"000" and W.Unit<>"999"

UNION SELECT
Sum(N.CPPTcnt)+Sum(N.OPSECcnt)+Sum(N.SafetyCnt),
Sum(N.CPPTcnt) ,
Sum(N.OPSECcnt),
Sum(N.SafetyCnt),
(select first(wing) from member)&"WG Totals on "&(SELECT FORMAT(DownLoadDate,"dd mmm yyyy") FROM DownLoadDate),
count(*)

FROM  (WingUnits AS W inner join SubUnitsWithCadets as C on W.Unit=C.unit) LEFT JOIN qryCadetCpptOPSECCountsSUB AS N ON N.Unit=W.Unit
Where W.Unit<>"000" and W.Unit<>"999")  AS X
ORDER BY Total DESC , UnitName;
