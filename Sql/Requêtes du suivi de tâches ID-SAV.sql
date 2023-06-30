SELECT u.Intitule Intervenant, i.NomClient NomClient, REPLACE(i.Description,'!','') as Description, id.Notes Notes,
	i.Code Code, i.PlanningDebut PlanningDebut, i.PlanningFin PlanningFin, i.TypePlanification TypePlanification, 
    i.FluxStatut FluxStatut, id.LibelleLong Objet, id.Description Detail
FROM sav.Interventions i
JOIN acl.Utilisateurs u on i.OperateurId = u.Id
LEFT JOIN sav.InterventionDetails id on id.InterventionId = i.Id
WHERE i.Description like '!%'
AND (i.TypePlanification != 3 OR i.FluxStatut != 3)
order by u.Intitule desc, i.Code, i.PlanningDebut