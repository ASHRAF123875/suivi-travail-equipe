# Version finale Excel : bouton d'alerte et dashboard

Le dépôt contient le module VBA `SuiviTravail.bas`.

## Installation

1. Ouvrez votre classeur Excel et enregistrez-le sous **Classeur Excel prenant en charge les macros (*.xlsm)**.
2. Préparez une feuille appelée `Suivi` avec ces en-têtes en ligne 1 :
   `Projet | Usine | N° article | Statut | Avancement (%) | Date de début | Date de fin | Jours restants | Date d'envoi email | Email envoyé | Alerte | Remarques | Email destinataire`
3. Appuyez sur `Alt+F11`, puis **Fichier > Importer un fichier** et sélectionnez `SuiviTravail.bas`.
4. Exécutez une seule fois la macro `InstallerClasseur`.
5. La feuille `Dashboard` est créée avec les KPI et le bouton **Actualiser les alertes**.
6. Utilisez le bouton pour recalculer les jours restants, les alertes, les couleurs et les indicateurs.

## Envoi des emails

La macro `EnvoyerAlertesOutlook` prépare dans Outlook les emails des lignes marquées `RETARD`, `URGENT` ou `EMAIL` lorsque `Email envoyé` n'est pas `Oui`.

Par sécurité, le code utilise `.Display` : vérifiez le message avant de l'envoyer. Après validation de vos tests, vous pouvez remplacer `.Display` par `.Send` dans le module.

Conditions requises : Outlook installé et compte configuré. Les macros doivent être autorisées dans Excel.

## Couleurs automatiques

- Rouge : retard
- Orange : urgent
- Jaune : email à traiter
- Vert : terminé
- Bleu : en cours
- Gris : en attente

## Important

Le module VBA ne peut pas être inclus dans un fichier `.xlsx` : utilisez impérativement le format `.xlsm`. Le fichier `.bas` est prêt à être importé dans Excel.
