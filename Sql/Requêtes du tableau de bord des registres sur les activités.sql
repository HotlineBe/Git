SELECT Code , TypePlanification, FluxStatut, OperateurId, DateOperation,
	   Description, CreatedOn, ModifiedOn, PlanningDebut, PlanningFin, NomClient 
FROM sav.Interventions i 
WHERE OperateurId in (36,4916,14)
AND DateOperation > '2023-01-01 10:52:13.905 -10:00'
AND DeletedOn IS NULL
ORDER BY DateOperation DESC 