# SUIVI DE TRAVAIL - VERSION PROFESSIONNELLE AVEC FORMATAGE ET COULEURS AUTOMATIQUES

## 📋 STRUCTURE COMPLÈTE DU FICHIER EXCEL

### ✨ Feuille 1 : SUIVI (Mise en forme avancée)

**Colonnes :**
- A : Projet
- B : Usine
- C : N° article
- D : Statut
- E : Avancement (%)
- F : Date de début
- G : Date de fin
- H : Jours restants
- I : Date d'envoi email
- J : Email envoyé
- K : Alerte
- L : Remarques
- M : Email destinataire

---

## 🎨 MISE EN FORME EXCEL - FORMATAGE AUTOMATIQUE

### 1️⃣ Format de l'en-tête (Ligne 1)

**Sélectionnez A1:M1**

Propriétés :
- Couleur de fond : Bleu foncé (#1F4E78)
- Couleur du texte : Blanc
- Gras : OUI
- Alignement : Centré
- Police : Calibri, 11 pt
- Hauteur de ligne : 25
- Bordure : Oui, épaisse

---

### 2️⃣ Format des données (Lignes 2 à 100)

**Colonne A - Projet :**
- Largeur : 120 px
- Alignement : Gauche
- Format : Texte standard

**Colonne B - Usine :**
- Largeur : 90 px
- Alignement : Gauche
- Format : Texte standard

**Colonne C - N° article :**
- Largeur : 100 px
- Alignement : Centre
- Format : Texte standard

**Colonne D - Statut :**
- Largeur : 110 px
- Alignement : Centre
- Format : Texte standard
- **Liste déroulante :** En attente, En cours, En retard, Terminé, Validé

**Colonne E - Avancement (%) :**
- Largeur : 100 px
- Alignement : Centre
- Format : Pourcentage (0%)
- Min : 0 | Max : 100

**Colonne F - Date de début :**
- Largeur : 100 px
- Alignement : Centre
- Format : Date (jj/mm/aaaa)

**Colonne G - Date de fin :**
- Largeur : 100 px
- Alignement : Centre
- Format : Date (jj/mm/aaaa)

**Colonne H - Jours restants :**
- Largeur : 95 px
- Alignement : Centre
- Format : Nombre
- **Formule :** `=SI(D2="Terminé";0;G2-AUJOURDHUI())`

**Colonne I - Date d'envoi email :**
- Largeur : 110 px
- Alignement : Centre
- Format : Date (jj/mm/aaaa)

**Colonne J - Email envoyé :**
- Largeur : 80 px
- Alignement : Centre
- Format : Texte
- **Liste déroulante :** Oui, Non

**Colonne K - Alerte :**
- Largeur : 100 px
- Alignement : Centre
- Format : Texte
- **Formule :** `=SI(ET(D2<>"Terminé";G2<AUJOURDHUI());"🔴 RETARD";SI(ET(E2<50;G2-AUJOURDHUI()<=3);"🟠 URGENT";SI(ET(I2<=AUJOURDHUI();J2="Non");"🟡 EMAIL";"")))`

**Colonne L - Remarques :**
- Largeur : 250 px
- Alignement : Gauche
- Format : Texte
- Retour à la ligne automatique : OUI

**Colonne M - Email destinataire :**
- Largeur : 200 px
- Alignement : Gauche
- Format : Texte

---

## 🎯 MISE EN FORME CONDITIONNELLE (COLORS AUTOMATIQUES)

### ✅ Sélectionnez les lignes de données : A2:M100

**Règle 1 - RETARD (Fond rouge clair)**
- Formule : `=ET(D2<>"Terminé";G2<AUJOURDHUI())`
- Couleur de fond : Rouge clair (#FFC7CE)
- Couleur du texte : Rouge foncé (#9C0006)
- Gras : OUI

**Règle 2 - URGENT (Fond orange clair)**
- Formule : `=ET(E2<50;G2-AUJOURDHUI()<=3)`
- Couleur de fond : Orange clair (#FFEB9C)
- Couleur du texte : Orange foncé (#9C6500)

**Règle 3 - EMAIL À ENVOYER (Fond jaune clair)**
- Formule : `=ET(I2<=AUJOURDHUI();J2="Non")`
- Couleur de fond : Jaune clair (#FFFFCC)
- Couleur du texte : Noir

**Règle 4 - TERMINÉ (Fond vert clair)**
- Formule : `=D2="Terminé"`
- Couleur de fond : Vert clair (#C6EFCE)
- Couleur du texte : Vert foncé (#006100)

**Règle 5 - EN COURS (Fond bleu clair)**
- Formule : `=D2="En cours"`
- Couleur de fond : Bleu clair (#DDEBF7)
- Couleur du texte : Bleu foncé (#002060)

**Règle 6 - EN ATTENTE (Fond gris clair)**
- Formule : `=D2="En attente"`
- Couleur de fond : Gris clair (#E7E6E6)
- Couleur du texte : Gris foncé (#595959)

**Règle 7 - VALIDÉ (Fond bleu très pâle)**
- Formule : `=D2="Validé"`
- Couleur de fond : Bleu très pâle (#F2F2F2)
- Couleur du texte : Noir

---

## 📊 Feuille 2 : DASHBOARD (Tableau de bord professionnel)

### En-tête du Dashboard

```
Cellule A1 : "TABLEAU DE BORD - SUIVI DE TRAVAIL"
- Fusion A1:D1
- Couleur de fond : Bleu foncé (#1F4E78)
- Couleur du texte : Blanc
- Gras : OUI
- Taille : 18 pt
- Hauteur de ligne : 35
- Alignement : Centre
```

### Indicateurs KPI

| | A | B |
|---|---|---|
| 3 | **Total travaux** | **=COUNTA(Suivi!A2:A1000)** |
| 4 | **En cours** | **=COUNTIF(Suivi!D2:D1000;"En cours")** |
| 5 | **Terminé** | **=COUNTIF(Suivi!D2:D1000;"Terminé")** |
| 6 | **En retard** | **=COUNTIF(Suivi!D2:D1000;"En retard")** |
| 7 | **Avancement moyen** | **=IFERROR(AVERAGE(Suivi!E2:E1000);"--")** |
| 8 | **Alertes à traiter** | **=COUNTIF(Suivi!K2:K1000;"🔴 RETARD")+COUNTIF(Suivi!K2:K1000;"🟠 URGENT")** |
| 9 | **Emails à envoyer** | **=COUNTIF(Suivi!K2:K1000;"🟡 EMAIL")** |

### Format des KPI

**Colonne A (Étiquettes) :**
- Couleur de fond : Bleu clair (#D9EAF7)
- Couleur du texte : Bleu foncé (#1F4E78)
- Gras : OUI
- Largeur : 200 px
- Bordure : OUI

**Colonne B (Valeurs) :**
- Couleur de fond : Jaune clair (#FFFFCC)
- Couleur du texte : Noir
- Gras : OUI
- Taille : 14 pt
- Largeur : 120 px
- Alignement : Centre
- Bordure : OUI

**Format colonne B7 :** Pourcentage (0%)

---

## 📈 Feuille 3 : DONNÉES (Listes déroulantes)

### Statuts disponibles

```
A1 : "Statut"
A2 : "En attente"
A3 : "En cours"
A4 : "En retard"
A5 : "Terminé"
A6 : "Validé"
```

### Usines disponibles

```
C1 : "Usine"
C2 : "Usine 1"
C3 : "Usine 2"
C4 : "Usine 3"
C5 : "Usine 4"
C6 : "Usine 5"
```

### Projets disponibles

```
E1 : "Projet"
E2 : "Projet Alpha"
E3 : "Projet Beta"
E4 : "Projet Gamma"
E5 : "Projet Delta"
E6 : "Projet Epsilon"
```

### Format de la feuille Données

- Couleur de fond : Gris très clair (#F2F2F2)
- Gras pour les en-têtes : OUI
- Bordure autour de chaque cellule : OUI

---

## 🔀 LISTES DÉROULANTES À CRÉER

### Sur la feuille SUIVI :

**Colonne A - Projet (A2:A100) :**
- Allez dans : Données > Validation des données
- Type : Liste
- Source : =Données!$E$2:$E$6
- Message d'erreur : "Veuillez sélectionner un projet valide"

**Colonne B - Usine (B2:B100) :**
- Source : =Données!$C$2:$C$6
- Message d'erreur : "Veuillez sélectionner une usine valide"

**Colonne D - Statut (D2:D100) :**
- Source : =Données!$A$2:$A$6
- Message d'erreur : "Veuillez sélectionner un statut valide"

**Colonne J - Email envoyé (J2:J100) :**
- Type : Liste
- Source : Oui;Non
- Message d'erreur : "Veuillez sélectionner Oui ou Non"

---

## 📊 DONNÉES D'EXEMPLE À COPIER

```csv
Projet Alpha,Usine 1,ART-001,En cours,60,01/09/2026,15/09/2026,,05/09/2026,Non,,Pièce en cours de fabrication,chef.alpha@entreprise.com
Projet Alpha,Usine 2,ART-002,Terminé,100,01/08/2026,10/08/2026,,02/08/2026,Oui,,Validation terminée,chef.alpha@entreprise.com
Projet Beta,Usine 1,ART-003,En attente,10,15/09/2026,30/09/2026,,08/09/2026,Non,,À confirmer avec le service,chef.beta@entreprise.com
Projet Gamma,Usine 3,ART-004,En retard,45,01/09/2026,20/09/2026,,04/09/2026,Non,🔴 RETARD,Retard dû à la disponibilité des matières,chef.gamma@entreprise.com
Projet Delta,Usine 2,ART-005,En cours,80,05/09/2026,18/09/2026,,06/09/2026,Non,,Contrôle qualité en cours,chef.delta@entreprise.com
Projet Delta,Usine 1,ART-006,Terminé,100,20/08/2026,29/08/2026,,21/08/2026,Oui,,Livraison effectuée,chef.delta@entreprise.com
Projet Epsilon,Usine 4,ART-007,En cours,30,12/09/2026,25/09/2026,,08/09/2026,Non,🟡 EMAIL,Révision des planches à revoir,chef.epsilon@entreprise.com
```

---

## 🔧 FILTRES ET TRI AUTOMATIQUES

**Sur la feuille SUIVI :**
- Sélectionnez l'en-tête (A1:M1)
- Allez dans : Données > Filtre automatique
- Les flèches de filtre apparaissent sur chaque colonne

---

## ✅ STYLE PROFESSIONNEL - COULEURS UTILISÉES

| Élément | Couleur HEX | Utilisation |
|---|---|---|
| En-têtes | #1F4E78 | Bleu foncé professionnel |
| Terminé | #C6EFCE | Vert clair (succès) |
| En cours | #DDEBF7 | Bleu clair (information) |
| En attente | #E7E6E6 | Gris clair (neutre) |
| En retard | #FFC7CE | Rouge clair (alerte) |
| Urgent | #FFEB9C | Orange clair (warning) |
| Email | #FFFFCC | Jaune clair (rappel) |
| KPI | #D9EAF7 | Bleu pâle (highlight) |

---

## 📱 RESPONSIVE DESIGN

**Largeurs optimales des colonnes :**
- Projet : 120 px
- Usine : 90 px
- N° article : 100 px
- Statut : 110 px
- Avancement : 100 px
- Dates : 100 px chacune
- Jours restants : 95 px
- Alerte : 100 px
- Remarques : 250 px
- Email : 200 px

---

## 🚀 ÉTAPES D'INSTALLATION

1. ✅ Créer feuille "Suivi" avec en-têtes
2. ✅ Appliquer formatage de l'en-tête (bleu foncé, blanc, gras)
3. ✅ Ajouter formules aux colonnes H et K
4. ✅ Créer listes déroulantes
5. ✅ Appliquer mise en forme conditionnelle (7 règles)
6. ✅ Créer feuille "Dashboard" avec KPI
7. ✅ Créer feuille "Données" avec listes
8. ✅ Ajouter filtre automatique
9. ✅ Copier données d'exemple
10. ✅ Tester les alertes en changeant les dates

---

## 💾 ENREGISTREMENT

- Format : **Excel (.xlsx)**
- Nom : **Suivi_Travail_Equipe_v2.xlsx**
- Emplacement : Bureau ou Dossier Partagé

---

## 📧 BONUS : INTÉGRATION EMAIL (Optionnel)

Pour ajouter l'envoi email automatique :
1. Ouvrez l'éditeur VBA (Alt+F11)
2. Insérez un module
3. Collez le code VBA fourni précédemment
4. Créez un bouton "Envoyer Emails"

---

**Version :** 2.5 Professionnel avec Formatage
**Date :** 30/09/2026
**Niveau de complexité :** Intermédiaire
**Temps d'installation :** 30-45 minutes
