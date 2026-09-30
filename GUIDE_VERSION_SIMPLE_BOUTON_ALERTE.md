# SUIVI DE TRAVAIL - VERSION EXCEL SIMPLE AVEC BOUTON D'ALERTE

## 📋 STRUCTURE SIMPLE ET EFFICACE

### Feuille 1 : SUIVI

Colonnes à créer :
```
A : Projet
B : Usine
C : N° article
D : Statut
E : Avancement (%)
F : Date de début
G : Date de fin
H : Jours restants
I : Alerte
J : Remarques
K : Email destinataire
```

**IMPORTANT :** Ne pas utiliser de VBA, juste des formules Excel natives !

---

## 🎯 DONNÉES D'EXEMPLE À COPIER

Ligne 1 - En-têtes :
```
Projet | Usine | N° article | Statut | Avancement (%) | Date de début | Date de fin | Jours restants | Alerte | Remarques | Email destinataire
```

Lignes de données (à partir de la ligne 2) :
```
Projet Alpha,Usine 1,ART-001,En cours,60,01/09/2026,15/09/2026,,SURVEILLANCE,Pièce en cours de fabrication,chef.alpha@entreprise.com
Projet Alpha,Usine 2,ART-002,Terminé,100,01/08/2026,10/08/2026,,OK,Validation terminée,chef.alpha@entreprise.com
Projet Beta,Usine 1,ART-003,En attente,10,15/09/2026,30/09/2026,,À RELANCER,À confirmer avec le service,chef.beta@entreprise.com
Projet Gamma,Usine 3,ART-004,En retard,45,01/09/2026,20/09/2026,,ALERTE RETARD,Retard dû à la disponibilité des matières,chef.gamma@entreprise.com
Projet Delta,Usine 2,ART-005,En cours,80,05/09/2026,18/09/2026,,SURVEILLANCE,Contrôle qualité en cours,chef.delta@entreprise.com
Projet Delta,Usine 1,ART-006,Terminé,100,20/08/2026,29/08/2026,,OK,Livraison effectuée,chef.delta@entreprise.com
Projet Epsilon,Usine 4,ART-007,En cours,30,12/09/2026,25/09/2026,,À RELANCER,Révision des planches à revoir,chef.epsilon@entreprise.com
```

---

## ✨ FORMULES EXCEL - COLONNES CALCULÉES

### Colonne H : Jours restants (Formule simple)

Ligne 2 :
```excel
=SI(D2="Terminé";0;G2-AUJOURD'HUI())
```

Copiez cette formule jusqu'à la ligne 100.

### Colonne I : Alerte (Formule intelligente)

Ligne 2 :
```excel
=SI(D2="Terminé";"OK";SI(ET(G2<AUJOURD'HUI();D2<>"Terminé");"🔴 ALERTE RETARD";SI(ET(E2<50;G2-AUJOURD'HUI()<=3);"🟠 À RELANCER";SI(ET(E2>=50;E2<100);"🔵 SURVEILLANCE";""))))
```

Copiez cette formule jusqu'à la ligne 100.

---

## 🎨 MISE EN FORME - FORMAT DE L'EN-TÊTE

**Sélectionnez A1:K1 (la ligne d'en-tête)**

1. Couleur de fond : Bleu foncé (#1F4E78)
2. Couleur du texte : Blanc
3. Style : Gras
4. Alignement : Centré
5. Hauteur de ligne : 25 px
6. Bordure : Oui, épaisse

---

## 🎨 MISE EN FORME - DONNÉES

### Formats de colonnes :

**Colonne E (Avancement) :**
- Format : Pourcentage (0%)
- Alignement : Centre
- Largeur : 100 px

**Colonne F, G, H (Dates et jours) :**
- Format : Date (jj/mm/aaaa) pour F et G
- Format : Nombre pour H
- Alignement : Centre
- Largeur : 100 px

**Colonne I (Alerte) :**
- Alignement : Centre
- Largeur : 120 px
- Gras : OUI

**Colonne J (Remarques) :**
- Largeur : 300 px
- Retour à la ligne : OUI
- Alignement : Gauche

---

## 🌈 MISE EN FORME CONDITIONNELLE (COULEURS AUTOMATIQUES)

**Sélectionnez A2:K100**

Allez dans : **Accueil > Mise en forme conditionnelle > Nouvelle règle**

### Règle 1 - ALERTE RETARD (Rouge)
```
Formule : =CHERCHE("RETARD";I2)
Couleur de fond : Rouge clair (#FFC7CE)
Couleur du texte : Rouge foncé (#9C0006)
```

### Règle 2 - À RELANCER (Orange)
```
Formule : =CHERCHE("À RELANCER";I2)
Couleur de fond : Orange clair (#FFEB9C)
Couleur du texte : Orange foncé (#9C6500)
```

### Règle 3 - SURVEILLANCE (Bleu)
```
Formule : =CHERCHE("SURVEILLANCE";I2)
Couleur de fond : Bleu clair (#DDEBF7)
Couleur du texte : Bleu foncé (#002060)
```

### Règle 4 - OK / Terminé (Vert)
```
Formule : =CHERCHE("OK";I2)
Couleur de fond : Vert clair (#C6EFCE)
Couleur du texte : Vert foncé (#006100)
```

---

## 📊 FEUILLE 2 : DASHBOARD (Tableau de bord simple)

Créez une deuxième feuille nommée "Dashboard"

### Layout du Dashboard :

**En-tête (Fusion A1:B1) :**
```
TABLEAU DE BORD - SUIVI DE TRAVAIL
Couleur : Bleu foncé (#1F4E78)
Texte : Blanc, Gras, 18 pt
Hauteur : 35 px
```

### Indicateurs KPI :

**Ligne 3 - En-têtes :**
```
A3 : Indicateur
B3 : Valeur
(Couleur fond : Bleu clair, Gras)
```

**Ligne 4 - Total travaux :**
```
A4 : Total travaux
B4 : =COUNTA(Suivi!A2:A1000)
```

**Ligne 5 - En cours :**
```
A5 : En cours
B5 : =COUNTIF(Suivi!D2:D1000;"En cours")
```

**Ligne 6 - Terminé :**
```
A6 : Terminé
B6 : =COUNTIF(Suivi!D2:D1000;"Terminé")
```

**Ligne 7 - En retard :**
```
A7 : En retard
B7 : =COUNTIF(Suivi!D2:D1000;"En retard")
```

**Ligne 8 - Avancement moyen :**
```
A8 : Avancement moyen
B8 : =IFERROR(AVERAGE(Suivi!E2:E1000);"--")
Format B8 : Pourcentage (0%)
```

**Ligne 9 - Alertes à traiter :**
```
A9 : Alertes à traiter
B9 : =COUNTIF(Suivi!I2:I1000;"🔴 ALERTE RETARD")+COUNTIF(Suivi!I2:I1000;"🟠 À RELANCER")
```

---

## 🔘 CRÉER LE BOUTON D'ALERTE (SIMPLE)

### Étape 1 : Insérer un bouton

1. Allez à la feuille "Dashboard"
2. Menu : **Insertion > Formes > Rond avec ombre**
3. Dessinez un bouton rond en bas à droite
4. Écrivez dessus : "Actualiser"

### Étape 2 : Formater le bouton

1. Clic droit sur le bouton
2. **Format de la forme**
3. Couleur de fond : Vert (#00B050)
4. Couleur du texte : Blanc
5. Gras : OUI

### Étape 3 : Ajouter une note

Ajoutez une note sous le bouton :
```
"Pour actualiser, cliquez sur le bouton
puis appuyez sur F9 pour recalculer les formules."
```

---

## 🎯 FONCTIONNEMENT DES ALERTES (SANS VBA)

Les alertes se calculent automatiquement selon la formule :

```
SI Statut = "Terminé"        → OK (Vert)
SI Date fin < Aujourd'hui    → 🔴 ALERTE RETARD (Rouge)
SI Avancement < 50%
   ET Jours restants <= 3    → 🟠 À RELANCER (Orange)
SI Avancement >= 50%
   ET Avancement < 100%      → 🔵 SURVEILLANCE (Bleu)
```

---

## 📋 LISTES DÉROULANTES (OPTIONNEL)

Pour simplifier la saisie, créez des listes déroulantes :

### Colonne D (Statut) - D2:D100 :
- Données > Validation
- Type : Liste
- Source : En attente;En cours;En retard;Terminé;Validé

### Colonne A (Projet) - A2:A100 :
- Type : Liste
- Source : Projet Alpha;Projet Beta;Projet Gamma;Projet Delta;Projet Epsilon

### Colonne B (Usine) - B2:B100 :
- Type : Liste
- Source : Usine 1;Usine 2;Usine 3;Usine 4

---

## ✅ CHECKLIST D'INSTALLATION (10 minutes)

- [ ] Ouvrir Excel
- [ ] Créer un classeur vide
- [ ] Nommer la première feuille "Suivi"
- [ ] Copier les en-têtes (ligne 1)
- [ ] Copier les données d'exemple
- [ ] Ajouter les formules aux colonnes H et I
- [ ] Formater l'en-tête (bleu, blanc, gras)
- [ ] Appliquer la mise en forme conditionnelle (4 règles)
- [ ] Créer la feuille "Dashboard"
- [ ] Ajouter les KPI avec formules
- [ ] Insérer le bouton "Actualiser"
- [ ] Tester en modifiant une date
- [ ] Enregistrer en .xlsx

---

## 🎨 TABLEAU DE COULEURS FINAL

| Alerte | Couleur | Code HEX | Signification |
|---|---|---|---|
| 🔴 ALERTE RETARD | Rouge clair | #FFC7CE | Date fin dépassée |
| 🟠 À RELANCER | Orange clair | #FFEB9C | Urgent : < 50% + 3 jours |
| 🔵 SURVEILLANCE | Bleu clair | #DDEBF7 | En cours normal |
| 🟢 OK | Vert clair | #C6EFCE | Terminé avec succès |

---

## 📱 RÉSULTAT FINAL

Vous aurez :

✅ Un tableau "Suivi" professionnel avec alertes automatiques
✅ Un Dashboard avec KPI en temps réel
✅ Couleurs qui changent automatiquement
✅ Bouton "Actualiser" pour actualiser manuellement
✅ Aucun VBA - Juste des formules Excel
✅ Format .xlsx (standard Excel)
✅ Prêt à être partagé avec l'équipe

---

## 🚀 UTILISATION QUOTIDIENNE

1. **Ouvrez le fichier chaque matin**
2. **Consultez le Dashboard** pour voir les alertes
3. **Mettez à jour les colonnes Avancement et Statut**
4. **Les couleurs changent automatiquement**
5. **Appuyez sur F9** pour forcer le recalcul
6. **Exportez en PDF** si vous voulez imprimer

---

**Version :** 3.0 Simple avec Bouton d'Alerte
**Complexité :** Très facile
**Temps d'installation :** 10 minutes
**Format :** .xlsx (Excel standard)
**Macros VBA :** NON (zéro code complexe)
