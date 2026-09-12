# 📊 Données d'Exemple

## Structure Complète du Classeur avec Exemples

---

## Feuil 1: Tableau de Bord

```
┌─────────────────────────────────────────────┐
│ TABLEAU DE BORD - GESTION RÉUNION           │
├─────────────────────────────────────────────┤
│ Nombre de membres actifs:        5          │
│ Total en caisse:                 1 045 000 FCFA
│ Dernier lieu de réunion:         Maison d'Aïssatou
│                                             │
│ ACTIONS:                                    │
│ [Prendre les présences]                     │
│ [Rapport Membre]                            │
│ [Clôture l'année]                           │
└─────────────────────────────────────────────┘
```

---

## Feuil 2: Membres

| Numéro membre | Nom      | Prénom    | Téléphone      | Date inscription |
|---------------|----------|-----------|----------------|------------------|
| 1             | KEITA    | Karim     | 06 11 22 33 44 | 15/01/2020       |
| 2             | DIALLO   | Mamadi    | 06 55 66 77 88 | 20/02/2020       |
| 3             | SOW      | Aïssatou  | 06 99 11 22 33 | 10/03/2020       |
| 4             | BA       | Ousmane   | 06 44 55 66 77 | 25/04/2020       |
| 5             | TOURE    | Fatima    | 06 88 99 00 11 | 05/05/2020       |

---

## Feuil 3: Réunions

| Date       | Lieu              | Total collectés | Numéro |
|------------|-------------------|-----------------|--------|
| 15/09/2026 | Maison de Karim   | 15 000          | 1      |
| 22/09/2026 | Maison de Mamadi  | 14 000          | 2      |
| 29/09/2026 | Maison d'Aïssatou | 13 000          | 3      |
| 06/10/2026 | Maison d'Ousmane  | 12 000          | 4      |
| 13/10/2026 | Maison de Fatima  | 11 000          | 5      |

---

## Feuil 4: Pointages (Exemple après 5 réunions)

| Date       | Nom et prénom      | Présences | Montant versé | Lieu              |
|------------|--------------------|-----------|---------------|-------------------|
| 15/09/2026 | KEITA Karim        | Oui       | 3000          | Maison de Karim   |
| 15/09/2026 | DIALLO Mamadi      | Oui       | 3000          | Maison de Karim   |
| 15/09/2026 | SOW Aïssatou       | Non       | 0             | Maison de Karim   |
| 15/09/2026 | BA Ousmane         | Oui       | 3000          | Maison de Karim   |
| 15/09/2026 | TOURE Fatima       | Oui       | 3000          | Maison de Karim   |
| 22/09/2026 | KEITA Karim        | Oui       | 2000          | Maison de Mamadi  |
| 22/09/2026 | DIALLO Mamadi      | Oui       | 3000          | Maison de Mamadi  |
| 22/09/2026 | SOW Aïssatou       | Oui       | 3000          | Maison de Mamadi  |
| 22/09/2026 | BA Ousmane         | Non       | 0             | Maison de Mamadi  |
| 22/09/2026 | TOURE Fatima       | Oui       | 3000          | Maison de Mamadi  |
| 29/09/2026 | KEITA Karim        | Oui       | 3000          | Maison d'Aïssatou |
| 29/09/2026 | DIALLO Mamadi      | Oui       | 3000          | Maison d'Aïssatou |
| 29/09/2026 | SOW Aïssatou       | Oui       | 3000          | Maison d'Aïssatou |
| 29/09/2026 | BA Ousmane         | Oui       | 3000          | Maison d'Aïssatou |
| 29/09/2026 | TOURE Fatima       | Non       | 0             | Maison d'Aïssatou |

---

## Feuil 5: Dépenses

| Date       | Motif                      | Montant |
|------------|---------------------------|---------|
| 20/09/2026 | Boissons réunion          | 5 000   |
| 25/09/2026 | Carburant transport       | 10 000  |
| 01/10/2026 | Collations réunion        | 8 000   |

---

## Feuil 6: Caisse

| Total caisse |
|--------------|
| 1 045 000    |

**Calcul**: 
- Réunion 1: 4 × 3000 = 12 000
- Réunion 2: 3 × 3000 + 1 × 2000 = 11 000
- Réunion 3: 4 × 3000 = 12 000
- **Total reçu**: ~35 000 FCFA

---

## Feuil 7: Archives (Après clôture année)

Contient tous les pointages précédents archivés (reste vide jusqu'à la première clôture).

---

## 📈 Rapports d'Exemple

### Rapport de KEITA Karim (après 3 réunions):
```
Membre: KEITA Karim

Nombre de réunion: 3/3
Nombre d'absences: 0

Total versé à la caisse: 8 000 FCFA
```

### Rapport de DIALLO Mamadi (après 3 réunions):
```
Membre: DIALLO Mamadi

Nombre de réunion: 3/3
Nombre d'absences: 0

Total versé à la caisse: 9 000 FCFA
```

### Rapport de SOW Aïssatou (après 3 réunions):
```
Membre: SOW Aïssatou

Nombre de réunion: 2/3
Nombre d'absences: 1

Total versé à la caisse: 6 000 FCFA
```

---

## 🎯 Comment remplir votre propre classeur

### Étape 1: Créer les feuilles
Créez 7 feuilles avec les noms exacts:
1. Tableau de bord
2. Membres
3. Réunions
4. Pointages
5. Dépenses
6. Caisse
7. Archives

### Étape 2: Ajouter les en-têtes
Copiez les en-têtes de chaque feuille ci-dessus dans la ligne 1.

### Étape 3: Ajouter vos données
- Membres réels
- Réunions réelles
- Pointages réels

### Étape 4: Importer le code VBA
Suivez le guide d'installation.

### Étape 5: Tester
Utilisez les 3 boutons pour vérifier que tout fonctionne.

---

**Dernière mise à jour:** 2026-09-12
