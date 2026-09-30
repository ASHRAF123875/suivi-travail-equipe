# Application de Suivi de Travail - Version Dashboard

## Description
Application Excel complète pour le suivi de travail d'équipe avec :
- Tableau de suivi détaillé
- Dashboard avec statistiques en temps réel
- Mise en forme conditionnelle
- Listes déroulantes pour faciliter la saisie
- Calculs automatiques

## Structure du fichier Excel

### Feuille 1 : Dashboard
Affiche les statistiques principales :
- **Total des travaux** : nombre total de projets
- **En cours** : nombre de travaux en cours
- **Terminés** : nombre de travaux terminés
- **En retard** : nombre de travaux en retard
- **Avancement moyen** : pourcentage moyen d'avancement
- **Graphique** : visualisation par statut

### Feuille 2 : Suivi
Tableau principal avec toutes les colonnes :
- Projet
- Usine
- N° article
- Statut (liste déroulante)
- Avancement (%)
- Date de début
- Date de fin
- Date d'envoi email
- Remarques

### Feuille 3 : Données
Listes déroulantes pour :
- Projets disponibles
- Usines disponibles
- Statuts possibles

## Statuts disponibles
- En attente
- En cours
- En retard
- Terminé
- Validé

## Mise en forme
Les lignes du tableau "Suivi" changent de couleur selon le statut :
- **Vert** : Terminé
- **Bleu** : En cours
- **Jaune** : En attente
- **Orange** : En retard
- **Gris** : Validé

## Formules principales

### Dans le Dashboard :
- `=COUNTA(Suivi!A2:A1000)` : Total des travaux
- `=COUNTIF(Suivi!D:D,"En cours")` : Nombre en cours
- `=COUNTIF(Suivi!D:D,"Terminé")` : Nombre terminés
- `=COUNTIF(Suivi!D:D,"En retard")` : Nombre en retard
- `=AVERAGE(Suivi!E:E)` : Avancement moyen

## Comment utiliser

### Ajouter un travail
1. Allez à la feuille "Suivi"
2. Cliquez sur une nouvelle ligne
3. Remplissez les champs :
   - Sélectionnez le Projet (liste déroulante)
   - Sélectionnez l'Usine (liste déroulante)
   - Entrez le N° article
   - Sélectionnez le Statut (liste déroulante)
   - Entrez l'Avancement en %
   - Entrez les dates (format JJ/MM/AAAA)
   - Entrez les remarques

### Filtrer les données
1. Cliquez sur l'en-tête du tableau
2. Utilisez les flèches de filtre pour trier par :
   - Projet
   - Usine
   - Statut
   - Date

### Consulter le Dashboard
1. Cliquez sur l'onglet "Dashboard"
2. Consultez les statistiques en temps réel
3. Les chiffres se mettent à jour automatiquement

## Conseils d'utilisation

✓ Mettez à jour régulièrement l'Avancement
✓ Changez le Statut quand le travail progresse
✓ Notez les remarques importantes
✓ Utilisez la Date d'envoi email pour tracer les communications
✓ Exportez en PDF pour les rapports

## Améliorations possibles
- Ajouter une colonne "Responsable"
- Ajouter une colonne "Priorité"
- Ajouter des alertes automatiques
- Créer un graphique de Gantt
- Ajouter un historique des modifications

---
Créé le : 30/09/2026
Version : 1.0
