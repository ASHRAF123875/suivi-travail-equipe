# SUIVI DE TRAVAIL - VERSION AVANCÉE AVEC VBA ET ENVOI EMAIL

## 📋 STRUCTURE DU FICHIER EXCEL

### Feuille 1 : SUIVI (Tableau principal)

| A | Projet | 
| B | Usine |
| C | N° article |
| D | Statut |
| E | Avancement (%) |
| F | Date de début |
| G | Date de fin |
| H | Jours restants |
| I | Date d'envoi email |
| J | Email envoyé |
| K | Alerte |
| L | Remarques |
| M | Email destinataire |

---

## 📊 FEUILLE SUIVI - FORMULES EXCEL

### Colonne H : Jours restants
```
=IF(D2="Terminé";0;G2-TODAY())
```

### Colonne I : Date d'envoi email (Automatique)
```
=IF(OR(D2="En cours";D2="En retard");TODAY();IF(D2="En attente";F2+7;""))
```

### Colonne J : Email envoyé (Dropdown)
- Créer une liste déroulante avec : Oui / Non
- Laisser vide par défaut (Non)

### Colonne K : Alerte (Mise en forme)
```
=IF(AND(D2<>"Terminé";G2<TODAY());"🔴 RETARD";IF(AND(E2<50;G2-TODAY()<=3);"🟠 URGENT";IF(AND(I2<=TODAY();J2="Non");"🟡 EMAIL";"")))
```

---

## 🎨 MISE EN FORME CONDITIONNELLE

**Sur la colonne K (Alerte) :**

1. **Retard (Rouge)**
   - Formule : `=AND(D2<>"Terminé";G2<TODAY())`
   - Couleur de fond : Rouge clair

2. **Urgent (Orange)**
   - Formule : `=AND(E2<50;G2-TODAY()<=3)`
   - Couleur de fond : Orange clair

3. **Email à envoyer (Jaune)**
   - Formule : `=AND(I2<=TODAY();J2="Non")`
   - Couleur de fond : Jaune clair

4. **Terminé (Vert)**
   - Formule : `=D2="Terminé"`
   - Couleur de fond : Vert clair

5. **En cours (Bleu)**
   - Formule : `=D2="En cours"`
   - Couleur de fond : Bleu clair

---

## 📧 DONNÉES - FEUILLE 3 : Adresses Email

Créez une feuille "Donnees" avec :

```
A1 : Projet | B1 : Email
A2 : Projet Alpha | B2 : chef.alpha@entreprise.com
A3 : Projet Beta | B3 : chef.beta@entreprise.com
A4 : Projet Gamma | B4 : chef.gamma@entreprise.com
A5 : Projet Delta | B5 : chef.delta@entreprise.com
A6 : Projet Epsilon | B6 : chef.epsilon@entreprise.com

D1 : Usine | E1 : Email
D2 : Usine 1 | E2 : usine1@entreprise.com
D3 : Usine 2 | E3 : usine2@entreprise.com
D4 : Usine 3 | E4 : usine3@entreprise.com
D5 : Usine 4 | E5 : usine4@entreprise.com

G1 : Statut | H1 : Couleur
G2 : En attente | H2 : Jaune
G3 : En cours | H3 : Bleu
G4 : En retard | H4 : Orange
G5 : Terminé | H5 : Vert
G6 : Validé | H6 : Gris
```

---

## 🤖 CODE VBA - ENVOI EMAIL AUTOMATIQUE

**Comment accéder à l'éditeur VBA :**
1. Ouvrez Excel
2. Appuyez sur `ALT + F11`
3. Allez dans : Insertion > Module
4. Collez le code ci-dessous

### CODE 1 : Envoyer un email simple

```vba
Sub EnvoyerEmailAlerte()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim ligne As Integer
    Dim projet As String
    Dim usine As String
    Dim article As String
    Dim statut As String
    Dim avancement As String
    Dim dateDebut As String
    Dim dateFin As String
    Dim remarques As String
    Dim email As String
    Dim sujet As String
    Dim corps As String

    Set ws = ThisWorkbook.Sheets("Suivi")
    Set OutApp = CreateObject("Outlook.Application")

    For ligne = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ' Récupérer les données
        projet = ws.Cells(ligne, 1).Value
        usine = ws.Cells(ligne, 2).Value
        article = ws.Cells(ligne, 3).Value
        statut = ws.Cells(ligne, 4).Value
        avancement = ws.Cells(ligne, 5).Value
        dateDebut = ws.Cells(ligne, 6).Value
        dateFin = ws.Cells(ligne, 7).Value
        remarques = ws.Cells(ligne, 12).Value
        email = ws.Cells(ligne, 13).Value

        ' Vérifier si email doit être envoyé
        If ws.Cells(ligne, 9).Value <= Now() And ws.Cells(ligne, 10).Value = "Non" And email <> "" Then
            Set OutMail = OutApp.CreateItem(0)
            
            ' Créer le sujet et le corps du mail
            sujet = "⚠️ ALERTE SUIVI - " & projet & " - " & article
            
            corps = "Bonjour," & vbCrLf & vbCrLf & _
                    "Voici un rappel concernant votre tâche :" & vbCrLf & vbCrLf & _
                    "📌 Projet : " & projet & vbCrLf & _
                    "🏭 Usine : " & usine & vbCrLf & _
                    "📦 N° Article : " & article & vbCrLf & _
                    "📊 Statut : " & statut & vbCrLf & _
                    "⏳ Avancement : " & avancement & "%" & vbCrLf & _
                    "📅 Date début : " & dateDebut & vbCrLf & _
                    "📅 Date fin : " & dateFin & vbCrLf & _
                    "📝 Remarques : " & remarques & vbCrLf & vbCrLf & _
                    "Merci de mettre à jour le statut dès que possible." & vbCrLf & vbCrLf & _
                    "Cordialement," & vbCrLf & _
                    "Système de suivi d'équipe"

            With OutMail
                .To = email
                .CC = ""
                .BCC = ""
                .Subject = sujet
                .Body = corps
                .Send ' Envoyer automatiquement
                ' .Display ' Afficher avant d'envoyer
            End With
            
            ' Marquer l'email comme envoyé
            ws.Cells(ligne, 10).Value = "Oui"
            ws.Cells(ligne, 10).Interior.Color = RGB(0, 255, 0)
        End If

    Next ligne

    MsgBox "Emails envoyés avec succès !", vbInformation, "Notification"
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
```

### CODE 2 : Envoyer un email avec affichage avant envoi

```vba
Sub EnvoyerEmailAvecConfirmation()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim ligne As Integer
    Dim projet As String
    Dim usine As String
    Dim article As String
    Dim statut As String
    Dim avancement As String
    Dim dateFin As String
    Dim remarques As String
    Dim email As String
    Dim sujet As String
    Dim corps As String
    Dim compteur As Integer

    Set ws = ThisWorkbook.Sheets("Suivi")
    Set OutApp = CreateObject("Outlook.Application")
    compteur = 0

    For ligne = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ' Récupérer les données
        projet = ws.Cells(ligne, 1).Value
        usine = ws.Cells(ligne, 2).Value
        article = ws.Cells(ligne, 3).Value
        statut = ws.Cells(ligne, 4).Value
        avancement = ws.Cells(ligne, 5).Value
        dateFin = ws.Cells(ligne, 7).Value
        remarques = ws.Cells(ligne, 12).Value
        email = ws.Cells(ligne, 13).Value

        ' Vérifier si email doit être envoyé
        If ws.Cells(ligne, 9).Value <= Now() And ws.Cells(ligne, 10).Value = "Non" And email <> "" Then
            Set OutMail = OutApp.CreateItem(0)
            
            sujet = "⚠️ ALERTE - " & projet & " - " & article
            
            corps = "Bonjour," & vbCrLf & vbCrLf & _
                    "Projet : " & projet & vbCrLf & _
                    "Usine : " & usine & vbCrLf & _
                    "N° Article : " & article & vbCrLf & _
                    "Statut : " & statut & vbCrLf & _
                    "Avancement : " & avancement & "%" & vbCrLf & _
                    "Date fin : " & dateFin & vbCrLf & _
                    "Remarques : " & remarques

            With OutMail
                .To = email
                .Subject = sujet
                .Body = corps
                .Display ' Afficher le mail avant envoi
            End With
            
            compteur = compteur + 1
        End If

    Next ligne

    MsgBox compteur & " email(s) prêt(s) à être envoyé(s).", vbInformation, "Notification"
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
```

---

## 🔘 CRÉER UN BOUTON POUR ENVOYER LES EMAILS

1. **Insérez un bouton :**
   - Allez dans : Insertion > Formes > Bouton
   - Dessinez un bouton sur la feuille
   - Nommez-le : "Envoyer Emails"

2. **Assignez une macro au bouton :**
   - Clic droit sur le bouton
   - Sélectionnez : "Assigner une macro"
   - Choisissez : `EnvoyerEmailAlerte`

3. **Colorez le bouton :**
   - Clic droit > Format de la forme
   - Couleur : Vert ou Orange

---

## 📊 FEUILLE 2 : DASHBOARD

Ajoutez ces statistiques :

| | A | B |
|---|---|---|
| 1 | **TABLEAU DE BORD** | |
| 3 | Total travaux | =COUNTA(Suivi!A2:A1000) |
| 4 | En cours | =COUNTIF(Suivi!D2:D1000;"En cours") |
| 5 | Terminé | =COUNTIF(Suivi!D2:D1000;"Terminé") |
| 6 | En retard | =COUNTIF(Suivi!D2:D1000;"En retard") |
| 7 | Avancement moyen | =IFERROR(AVERAGE(Suivi!E2:E1000);0) |
| 8 | Alertes à traiter | =COUNTIF(Suivi!K2:K1000;"🔴 RETARD")+COUNTIF(Suivi!K2:K1000;"🟠 URGENT") |
| 9 | Emails à envoyer | =COUNTIF(Suivi!K2:K1000;"🟡 EMAIL") |

**Format B7 :** Pourcentage

---

## 📝 DONNÉES EXEMPLE POUR LA FEUILLE SUIVI

```
Projet Alpha | Usine 1 | ART-001 | En cours | 60 | 01/09/2026 | 15/09/2026 | 05/09/2026 | Non | | chef.alpha@entreprise.com | Pièce en cours
Projet Alpha | Usine 2 | ART-002 | Terminé | 100 | 01/08/2026 | 10/08/2026 | 02/08/2026 | Oui | | chef.alpha@entreprise.com | Validation OK
Projet Beta | Usine 1 | ART-003 | En attente | 10 | 15/09/2026 | 30/09/2026 | 08/09/2026 | Non | | chef.beta@entreprise.com | À confirmer
Projet Gamma | Usine 3 | ART-004 | En retard | 45 | 01/09/2026 | 20/09/2026 | 04/09/2026 | Non | 🔴 | chef.gamma@entreprise.com | Retard matière
Projet Delta | Usine 2 | ART-005 | En cours | 80 | 05/09/2026 | 18/09/2026 | 06/09/2026 | Non | | chef.delta@entreprise.com | Contrôle en cours
Projet Epsilon | Usine 4 | ART-007 | En cours | 30 | 12/09/2026 | 25/09/2026 | 08/09/2026 | Non | 🟡 | chef.epsilon@entreprise.com | À réviser
```

---

## ✅ CHECKLIST D'INSTALLATION

- [ ] Créer feuille "Suivi" avec toutes les colonnes
- [ ] Ajouter les formules aux colonnes H, I, K
- [ ] Ajouter la mise en forme conditionnelle
- [ ] Créer feuille "Donnees" avec listes
- [ ] Ajouter listes déroulantes (Statut, Projet, Usine)
- [ ] Créer feuille "Dashboard"
- [ ] Ajouter les formules du Dashboard
- [ ] Accéder à l'éditeur VBA (ALT+F11)
- [ ] Insérer les codes VBA
- [ ] Créer un bouton "Envoyer Emails"
- [ ] Assigner la macro au bouton
- [ ] Tester l'envoi d'email
- [ ] Ajouter les emails destinataires dans la colonne M

---

## 🔧 CONFIGURATION OUTLOOK

**Important :** Pour que le VBA fonctionne, vous devez avoir :
- Microsoft Outlook installé
- Un compte email configuré dans Outlook
- Les autorisations VBA activées dans Excel

**Activation des macros :**
1. Fichier > Options > Centre de gestion de la confidentialité
2. Centre de gestion de la confidentialité > Paramètres des macros
3. Sélectionnez : "Activer toutes les macros"

---

## 📧 FONCTIONNALITÉS

✅ Envoi automatique d'emails selon la date prévue
✅ Marquage automatique "Email envoyé" après envoi
✅ Alertes visuelles (couleurs et emojis)
✅ Calcul automatique des jours restants
✅ Dashboard en temps réel
✅ Listes déroulantes pour éviter les erreurs
✅ Formules de validation automatique
✅ Bouton unique pour envoyer tous les emails

---

## 🚀 UTILISATION AU QUOTIDIEN

1. **Chaque matin :**
   - Ouvrez le fichier Excel
   - Consultez le Dashboard pour voir les alertes
   - Cliquez sur "Envoyer Emails" pour notifier l'équipe

2. **Chaque jour :**
   - Mettez à jour les colonnes "Avancement" et "Statut"
   - Les alertes s'affichent automatiquement
   - Les emails se marquent comme envoyés

3. **Chaque semaine :**
   - Générez un rapport du Dashboard
   - Exportez en PDF si nécessaire

---

## ⚠️ NOTES IMPORTANTES

- Le VBA nécessite **Outlook** installé
- Les emails doivent être saisis dans la colonne M
- La date d'envoi se calcule automatiquement
- Les alertes s'affichent en temps réel
- Sauvegardez en format **.xlsm** (pas .xlsx)

---

**Version :** 2.0 Avancée
**Date :** 30/09/2026
**Support :** Equipe IT
