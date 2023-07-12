select * from (
	SELECT u.Intitule Intervenant, i.NomClient NomClient, REPLACE(i.Description,'!','') as Description, id.Notes Notes,
			i.Code Code, i.PlanningDebut PlanningDebut, i.PlanningFin PlanningFin, i.TypePlanification TypePlanification, 
		    i.FluxStatut FluxStatut, id.LibelleLong Objet, id.Description Detail
	FROM sav.Interventions i
	JOIN acl.Utilisateurs u on i.OperateurId = u.Id
	LEFT JOIN sav.InterventionDetails id on id.InterventionId = i.Id
	WHERE u.Intitule in ('Gilles', 'Stéphanie', 'Leilanie')
	AND convert(varchar, i.PlanningDebut, 23) = convert(varchar, getdate(), 23)
	UNION ALL 
	SELECT u.Intitule Intervenant, i.NomClient NomClient, REPLACE(i.Description,'!','') as Description, id.Notes Notes,
			i.Code Code, i.PlanningDebut PlanningDebut, i.PlanningFin PlanningFin, i.TypePlanification TypePlanification, 
		    i.FluxStatut FluxStatut, id.LibelleLong Objet, id.Description Detail
	FROM sav.Interventions i
	JOIN acl.Utilisateurs u on i.OperateurId = u.Id
	LEFT JOIN sav.InterventionDetails id on id.InterventionId = i.Id
	WHERE u.Intitule in ('Gilles', 'Stéphanie', 'Leilanie')
	AND convert(varchar, i.PlanningFin, 23) = convert(varchar, (getdate()+1), 23)
) t 
ORDER BY t.Intervenant