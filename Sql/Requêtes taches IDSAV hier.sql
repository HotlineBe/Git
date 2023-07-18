select distinct * from (
	SELECT u.Intitule Intervenant, i.NomClient NomClient, REPLACE(i.Description,'!','') as Description, id.Notes Notes,
			i.Code Code, i.PlanningDebut PlanningDebut, i.PlanningFin PlanningFin,
		    CASE i.TypePlanification
		    	WHEN 0 THEN 'Non planifié'
		        WHEN 1 THEN 'A planifier'
		        WHEN 2 THEN 'En cours'
		        WHEN 3 THEN 'Terminé'
		    END AS TypePlanification,
		    CASE i.FluxStatut
		    	WHEN 0 THEN 'Nouveau'
		        WHEN 1 THEN 'En cours'
		        WHEN 3 THEN 'Fermé'
		    END AS FluxStatut,
		    id.LibelleLong Objet, id.Description Detail
	FROM sav.Interventions i
	JOIN acl.Utilisateurs u on i.OperateurId = u.Id
	LEFT JOIN sav.InterventionDetails id on id.InterventionId = i.Id
	WHERE u.Intitule in ('Gilles', 'Stéphanie', 'Leilanie')
	AND convert(varchar, i.PlanningDebut, 23) = convert(varchar, (getdate()-1), 23)
	UNION ALL 
	SELECT u.Intitule Intervenant, i.NomClient NomClient, REPLACE(i.Description,'!','') as Description, id.Notes Notes,
			i.Code Code, i.PlanningDebut PlanningDebut, i.PlanningFin PlanningFin,
		    CASE i.TypePlanification
		    	WHEN 0 THEN 'Non planifié'
		        WHEN 1 THEN 'A planifier'
		        WHEN 2 THEN 'En cours'
		        WHEN 3 THEN 'Terminé'
		    END AS TypePlanification,
		    CASE i.FluxStatut
		    	WHEN 0 THEN 'Nouveau'
		        WHEN 1 THEN 'En cours'
		        WHEN 3 THEN 'Fermé'
		    END AS FluxStatut,
		    id.LibelleLong Objet, id.Description Detail
	FROM sav.Interventions i
	JOIN acl.Utilisateurs u on i.OperateurId = u.Id
	LEFT JOIN sav.InterventionDetails id on id.InterventionId = i.Id
	WHERE u.Intitule in ('Gilles', 'Stéphanie', 'Leilanie')
	AND convert(varchar, i.PlanningFin, 23) = convert(varchar, getdate()-1, 23)
) t 
ORDER BY t.Intervenant