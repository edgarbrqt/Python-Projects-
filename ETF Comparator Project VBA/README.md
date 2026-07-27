# ETF Comparator Tool — VBA

Outil Excel/VBA d'analyse comparative multi-ETF, développé en préparation d'un stage Sales ETF (Xtrackers/DWS).

## Fonctionnalités

- Calcul automatique de la **tracking difference** et de la **tracking error** pour plusieurs ETF, à partir de leurs rendements quotidiens et de ceux de leur indice de référence
- Calcul de la **performance cumulée** jour par jour
- Génération d'un tableau récapitulatif formaté (bordures, en-tête stylé)
- Mise en forme conditionnelle sur le TER (vert/orange selon un seuil)

## Concepts financiers appliqués

- **Tracking difference** : écart de performance cumulée entre l'ETF et son indice sur la période
- **Tracking error** : volatilité (écart-type annualisé) des écarts quotidiens de rendement — mesure la régularité du suivi, indépendamment du niveau de sur/sous-performance
- Ces deux métriques sont calculées *from scratch* en VBA (pas via une fonction Excel native), pour démontrer la compréhension du calcul sous-jacent

## Structure du fichier

- `Info_ETF` : liste des ETF, TER, et résultats du comparateur
- `Rendements` : historique de rendements quotidiens ETF/indice sur ~3 mois, pour 3 ETF (MSCI World, Euro Stoxx 50, MSCI Emerging Markets)

## Compétences techniques démontrées

- VBA : Functions et Subs, boucles, gestion multi-feuilles, arrays, formatage conditionnel
- Structuration du code en procédures indépendantes et réutilisables (calcul, formatage, orchestration via un `Sub Principal`)

## Utilisation

1. Ouvrir `ETF Comparator Project with VBA` (macros à activer)
2. Développeur → Macros → lancer `Principal`


---
*Projet réalisé en autodidacte (VBA) dans le cadre d'une préparation à un entretien de stage en Sales ETF chez DWS (Xtrackers).*
