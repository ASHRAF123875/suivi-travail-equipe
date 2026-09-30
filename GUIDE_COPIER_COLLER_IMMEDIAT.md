# VERSION EXCEL PRÊTE À COPIER-COLLER

## 📋 ÉTAPE 1 : COPIER-COLLER LE TABLEAU SUIVI

Ouvrez un nouveau fichier Excel et nommez la première feuille **"Suivi"**

### Copiez exactement ce texte dans Excel (à partir de A1) :

```
Projet	Usine	N° article	Statut	Avancement (%)	Date de début	Date de fin	Jours restants	Alerte	Remarques	Email destinataire
Projet Alpha	Usine 1	ART-001	En cours	60%	01/09/2026	15/09/2026		SURVEILLANCE	Pièce en cours de fabrication	chef.alpha@entreprise.com
Projet Alpha	Usine 2	ART-002	Terminé	100%	01/08/2026	10/08/2026		OK	Validation terminée	chef.alpha@entreprise.com
Projet Beta	Usine 1	ART-003	En attente	10%	15/09/2026	30/09/2026		À RELANCER	À confirmer avec le service	chef.beta@entreprise.com
Projet Gamma	Usine 3	ART-004	En retard	45%	01/09/2026	20/09/2026		ALERTE RETARD	Retard dû à la disponibilité des matières	chef.gamma@entreprise.com
Projet Delta	Usine 2	ART-005	En cours	80%	05/09/2026	18/09/2026		SURVEILLANCE	Contrôle qualité en cours	chef.delta@entreprise.com
Projet Delta	Usine 1	ART-006	Terminé	100%	20/08/2026	29/08/2026		OK	Livraison effectuée	chef.delta@entreprise.com
Projet Epsilon	Usine 4	ART-007	En cours	30%	12/09/2026	25/09/2026		À RELANCER	Révision des planches à revoir	chef.epsilon@entreprise.com
```

**Instructions de copie :**
1. Sélectionnez le texte ci-dessus
2. Copiez-le (Ctrl+C)
3. Ouvrez Excel
4. Cliquez sur la cellule A1
5. Collez (Ctrl+V)

---

## 🔧 ÉTAPE 2 : AJOUTER LES FORMULES

### Formule pour la colonne H (Jours restants)

Cliquez sur la cellule **H2** et entrez cette formule :

```excel
=SI(D2="Terminé";0;G2-AUJOURD'HUI())
```

Puis :
1. Appuyez sur **Entrée**
2. Sélectionnez H2
3. Copiez la cellule (Ctrl+C)
4. Sélectionnez la plage **H3:H100**
5. Collez (Ctrl+V)

---

### Formule pour la colonne I (Alerte - TRÈS IMPORTANTE)

Cliquez sur la cellule **I2** et entrez cette formule exactement :

```excel
=SI(D2="Terminé";"OK";SI(ET(G2<AUJOURD'HUI();D2<>"Terminé");"ALERTE RETARD";SI(ET(E2<0.5;G2-AUJOURD'HUI()<=3);"À RELANCER";SI(ET(E2>=0.5;E2<1);"SURVEILLANCE";""))))
```

Puis :
1. Appuyez sur **Entrée**
2. Sélectionnez I2
3. Copiez la cellule (Ctrl+C)
4. Sélectionnez la plage **I3:I100**
5. Collez (Ctrl+V)

---

## 🎨 ÉTAPE 3 : FORMATER L'EN-TÊTE

1. **Sélectionnez la ligne 1** (A1:K1)
2. Clic droit > **Format de cellule**
3. Onglet **Remplissage** > Couleur : Bleu (#1F4E78)
4. Onglet **Police** > Couleur : Blanc
5. Onglet **Alignement** > Alignement horizontal : Centré
6. Cliquez sur le bouton **Gras** dans la barre d'outils
7. Augmentez la hauteur de la ligne 1 à 25 px

---

## 🌈 ÉTAPE 4 : MISE EN FORME CONDITIONNELLE (COULEURS AUTOMATIQUES)

### Sélectionnez la plage A2:K100

1. Allez dans l'onglet **Accueil**
2. Cliquez sur **Mise en forme conditionnelle** > **Nouvelle règle**
3. Sélectionnez **"Utiliser une formule pour déterminer..."**

### Règle 1 : ALERTE RETARD (Rouge)

**Formule :**
```
=CHERCHE("RETARD";I2)
```

**Couleur de fond :** Rouge clair (#FFC7CE)
**Couleur du texte :** Rouge foncé (#9C0006)

### Règle 2 : À RELANCER (Orange)

**Formule :**
```
=CHERCHE("RELANCER";I2)
```

**Couleur de fond :** Orange clair (#FFEB9C)
**Couleur du texte :** Orange foncé (#9C6500)

### Règle 3 : SURVEILLANCE (Bleu)

**Formule :**
```
=CHERCHE("SURVEILLANCE";I2)
```

**Couleur de fond :** Bleu clair (#DDEBF7)
**Couleur du texte :** Bleu foncé (#002060)

### Règle 4 : OK (Vert)

**Formule :**
```
=CHERCHE("OK";I2)
```

**Couleur de fond :** Vert clair (#C6EFCE)
**Couleur du texte :** Vert foncé (#006100)

---

## 📊 ÉTAPE 5 : CRÉER LA FEUILLE DASHBOARD

1. Clic droit sur l'onglet "Suivi" en bas
2. Sélectionnez **Insérer une feuille**
3. Nommez-la : **Dashboard**

### Copiez ce contenu dans la feuille Dashboard :

**Cellule A1:**
```
TABLEAU DE BORD - SUIVI DE TRAVAIL
```

Formatage A1 :
- Fusion avec B1
- Couleur fond : Bleu (#1F4E78)
- Couleur texte : Blanc
- Gras : OUI
- Taille : 18 pt
- Hauteur : 35 px

---

### Tableau des KPI :

**À partir de la cellule A3, copiez :**

```
Indicateur	Valeur
Total travaux	
En cours	
Terminé	
En retard	
Avancement moyen	
Alertes à traiter	
```

---

### Formules du Dashboard :

**Cliquez sur B4** et entrez :
```excel
=COUNTA(Suivi!A2:A1000)
```

**Cliquez sur B5** et entrez :
```excel
=COUNTIF(Suivi!D2:D1000;"En cours")
```

**Cliquez sur B6** et entrez :
```excel
=COUNTIF(Suivi!D2:D1000;"Terminé")
```

**Cliquez sur B7** et entrez :
```excel
=COUNTIF(Suivi!D2:D1000;"En retard")
```

**Cliquez sur B8** et entrez :
```excel
=IFERROR(AVERAGE(Suivi!E2:E1000);"--")
```

Puis formatez B8 en pourcentage (clic droit > Format de cellule > Pourcentage)

**Cliquez sur B9** et entrez :
```excel
=COUNTIF(Suivi!I2:I1000;"ALERTE RETARD")+COUNTIF(Suivi!I2:I1000;"À RELANCER")
```

---

### Formatage des cellules B4:B9 :

1. Sélectionnez **B4:B9**
2. Clic droit > **Format de cellule**
3. Onglet **Remplissage** > Couleur : Jaune (#FFFFCC)
4. Onglet **Police** > Taille : 14 pt, Gras : OUI
5. Onglet **Alignement** > Alignement horizontal : Centré

---

## 🔘 ÉTAPE 6 : INSÉRER LE BOUTON D'ALERTE

1. Restez sur la feuille **Dashboard**
2. Menu **Insertion** > **Formes** > Sélectionnez un rond avec ombre
3. Dessinez un bouton rond en bas à gauche
4. Écrivez dessus : **"Actualiser"**
5. Clic droit sur le bouton > **Format de la forme**
6. Couleur de fond : Vert (#00B050)
7. Couleur du texte : Blanc
8. Gras : OUI

---

## ✅ ÉTAPE 7 : AJOUTER LES LISTES DÉROULANTES (OPTIONNEL)

### Pour la colonne D (Statut) :

1. Sélectionnez **D2:D100**
2. Menu **Données** > **Validation des données**
3. Type : **Liste**
4. Source : **En attente,En cours,En retard,Terminé,Validé**
5. Cliquez **OK**

### Pour la colonne A (Projet) :

1. Sélectionnez **A2:A100**
2. Menu **Données** > **Validation des données**
3. Type : **Liste**
4. Source : **Projet Alpha,Projet Beta,Projet Gamma,Projet Delta,Projet Epsilon**
5. Cliquez **OK**

### Pour la colonne B (Usine) :

1. Sélectionnez **B2:B100**
2. Menu **Données** > **Validation des données**
3. Type : **Liste**
4. Source : **Usine 1,Usine 2,Usine 3,Usine 4**
5. Cliquez **OK**

---

## 📱 ÉTAPE 8 : AJOUTER LE FILTRE AUTOMATIQUE

1. Sélectionnez la ligne 1 (A1:K1)
2. Menu **Données** > **Filtre automatique**
3. Des flèches apparaîtront sur chaque en-tête

---

## 💾 ÉTAPE 9 : ENREGISTRER LE FICHIER

1. Menu **Fichier** > **Enregistrer sous**
2. Nom : **Suivi_Travail_Equipe.xlsx**
3. Format : **Excel (.xlsx)**
4. Emplacement : Bureau ou dossier partagé
5. Cliquez **Enregistrer**

---

## ✨ RÉSULTAT FINAL

Vous devez avoir :

✅ **Feuille "Suivi"** avec :
   - Tableau professionnel avec en-têtes bleus
   - Données d'exemple
   - Formules automatiques pour Jours restants et Alertes
   - Couleurs qui changent selon l'alerte
   - Filtre automatique

✅ **Feuille "Dashboard"** avec :
   - KPI en temps réel
   - Bouton "Actualiser"
   - Résumé des alertes

✅ **Format .xlsx** prêt à être partagé

---

## 🎯 TESTER LE FICHIER

1. Ouvrez la feuille "Suivi"
2. Modifiez une date de fin pour qu'elle soit dans le passé
3. Appuyez sur **F9** pour recalculer les formules
4. La ligne doit devenir **rouge** avec "ALERTE RETARD"
5. Le Dashboard doit se mettre à jour automatiquement

---

## 📋 CHECKLIST DE VÉRIFICATION

- [ ] Tableau "Suivi" créé avec les 11 colonnes
- [ ] 7 lignes d'exemple ajoutées
- [ ] Formules H2 et I2 en place et copiées
- [ ] En-tête formaté (bleu, blanc, gras)
- [ ] Mise en forme conditionnelle appliquée (4 règles)
- [ ] Feuille "Dashboard" créée
- [ ] KPI avec formules ajoutés
- [ ] Bouton "Actualiser" inséré
- [ ] Listes déroulantes créées (optionnel)
- [ ] Filtre automatique activé
- [ ] Fichier enregistré en .xlsx
- [ ] Test des alertes effectué

---

## 🚀 UTILISATION AU QUOTIDIEN

**Chaque matin :**
1. Ouvrez le fichier
2. Consultez le Dashboard pour les alertes
3. Mettez à jour la colonne "Avancement"
4. Changez le statut si nécessaire
5. Appuyez sur **F9** pour recalculer
6. Les couleurs changent automatiquement

**Chaque semaine :**
- Exportez le Dashboard en PDF
- Partagez avec l'équipe

---

**VERSION :** 4.0 - Prête à copier-coller
**COMPLEXITÉ :** Très facile
**TEMPS :** 20 minutes
**VBA REQUIS :** NON
**OUTLOOK REQUIS :** NON
