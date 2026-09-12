# 📖 Guide d'Installation et d'Utilisation

## 🚀 Installation

### Étape 1: Créer le fichier Excel
1. Ouvrez **Microsoft Excel**
2. Créez un nouveau classeur
3. Renommez les feuilles comme suit:
   - **Feuil1** → "Tableau de bord"
   - **Feuil2** → "Membres"
   - **Feuil3** → "Réunions"
   - **Feuil4** → "Pointages"
   - **Feuil5** → "Dépenses"
   - **Feuil6** → "Caisse"
   - Créez une **Feuil7** → "Archives"

### Étape 2: Ajouter les en-têtes (Colonnes)

#### Feuil "Membres" (Ligne 1)
```
A1: Numéro membre
B1: Nom
C1: Prénom
D1: Téléphone
E1: Date inscription
```

#### Feuil "Réunions" (Ligne 1)
```
A1: Date
B1: Lieu
C1: Total collectés
D1: Numéro
```

#### Feuil "Pointages" (Ligne 1)
```
A1: Date
B1: Nom et prénom
C1: Présences
D1: Montant versé
E1: Lieu
```

#### Feuil "Dépenses" (Ligne 1)
```
A1: Date
B1: Motif
C1: Montant
```

#### Feuil "Caisse" (Ligne 1)
```
A1: Total caisse
```

#### Feuil "Archives" (Ligne 1)
```
A1: Date
B1: Nom et prénom
C1: Présences
D1: Montant versé
E1: Lieu
```

### Étape 3: Importer le code VBA

1. **Ouvrez l'Éditeur VBA**: `Alt + F11`
2. **Créez un nouveau module**:
   - Clic droit sur "VBAProject"
   - Sélectionnez "Insérer" → "Module"
3. **Copiez tout le code** du fichier `Module_Principal.vb`
4. **Collez-le** dans le module
5. **Sauvegardez**: `Ctrl + S`

### Étape 4: Créer les boutons

1. **Allez à la feuille "Tableau de bord"**
2. **Insérez 3 boutons** (Onglet Développeur → Insérer → Bouton):

#### Bouton 1: "Prendre les présences"
- Position: A9
- Assignez la macro: `BoutonPrendrePresences`
- Couleur: Bleu

#### Bouton 2: "Rapport Membre"
- Position: A11
- Assignez la macro: `BoutonRapportMembre`
- Couleur: Vert

#### Bouton 3: "Clôture l'année"
- Position: A13
- Assignez la macro: `BoutonClotureAnnee`
- Couleur: Rouge

**Comment assigner une macro à un bouton:**
1. Clic droit sur le bouton
2. Sélectionnez "Assigner une macro"
3. Choisissez le nom de la macro

### Étape 5: Ajouter des données de test

#### Exemple - Feuil "Membres":
```
Numéro membre | Nom    | Prénom | Téléphone   | Date inscription
1             | KEITA  | Karim  | 06XXXXXXXX  | 15/01/2020
2             | DIALLO | Mamadi | 06XXXXXXXX  | 20/02/2020
3             | SOW    | Aïssatou| 06XXXXXXXX | 10/03/2020
```

#### Exemple - Feuil "Réunions":
```
Date       | Lieu              | Total collectés | Numéro
15/09/2026 | Maison de Karim   | 9000           | 1
22/09/2026 | Maison de Mamadi  | 9000           | 2
29/09/2026 | Maison d'Aïssatou | 8000           | 3
```

---

## 📋 Utilisation

### 1️⃣ Tableau de Bord

**Vue d'ensemble** avec:
- ✅ Nombre de membres actifs
- ✅ Total en caisse
- ✅ Dernier lieu de réunion
- ✅ 3 boutons d'actions

**Mise à jour automatique** après chaque action.

---

### 2️⃣ Bouton: "Prendre les présences"

**Étapes:**

1. **Cliquez sur le bouton** "Prendre les présences"

2. **Sélectionnez une réunion**:
   ```
   Liste des réunions:
   1. 15/09/2026 - Maison de Karim
   2. 22/09/2026 - Maison de Mamadi
   3. 29/09/2026 - Maison d'Aïssatou
   
   Tapez: 1 (ou 2 ou 3)
   ```

3. **Pour chaque membre**, répondez aux questions:
   ```
   Q1: Est-ce que KEITA Karim était présent?
   Réponse: Oui (ou Non)
   
   Q2: Montant versé par KEITA Karim (FCFA):
   Réponse: 3000 (par défaut) ou 1000 ou 2000 (acomptes)
   ```

4. **Les données s'enregistrent automatiquement** et le total en caisse se met à jour.

**Format des présences**: `Oui` ou `Non` (exact)

**Montants acceptés**:
- 3000 FCFA (montant standard)
- 1000 FCFA (acompte)
- 2000 FCFA (acompte)
- Autre montant custom (accepté)

---

### 3️⃣ Bouton: "Rapport Membre"

**Affiche les statistiques** d'un membre:

1. **Cliquez sur le bouton** "Rapport Membre"

2. **Entrez le nom**:
   ```
   Tapez: KEITA
   ```

3. **Entrez le prénom**:
   ```
   Tapez: Karim
   ```

4. **Résultat affiché**:
   ```
   Membre: KEITA Karim
   
   Nombre de réunion: 66/83
   Nombre d'absences: 17
   
   Total versé à la caisse: 198 000 FCFA
   ```

**Explication**:
- `66/83` = 66 réunions auxquelles il a participé sur 83 réunions au total
- `17` = Nombre de fois où il était absent
- `198 000 FCFA` = Somme totale qu'il a versée

---

### 4️⃣ Bouton: "Clôture l'année"

**Archive tous les pointages** et réinitialise:

1. **Cliquez sur le bouton** "Clôture l'année"

2. **Confirmez** la clôture:
   ```
   Êtes-vous sûr de vouloir clôturer l'année?
   Tous les pointages seront archivés.
   
   [Oui] [Non]
   ```

3. **Résultat**:
   - ✅ Tous les pointages sont transférés à la feuille "Archives"
   - ✅ La feuille "Pointages" est vidée (prête pour la nouvelle année)
   - ✅ La liste des "Réunions" reste intacte
   - ✅ La "Caisse" reste intacte

4. **Vérification**: Allez dans "Archives" pour voir tous les pointages précédents

---

## 🔧 Dépannage

### Problème: Les boutons ne fonctionnent pas
**Solution:**
1. Allez dans `Fichier` → `Options` → `Centre de gestion de la confidentialité`
2. Activez les **macros**
3. Redémarrez Excel

### Problème: Le rapport affiche "0 FCFA"
**Solution:**
1. Vérifiez que le **nom et prénom** correspondent exactement
2. Vérifiez que la feuille "Pointages" contient des données
3. Assurez-vous que les montants sont en **nombres** (pas texte)

### Problème: La caisse n'augmente pas
**Solution:**
1. Vérifiez que vous avez **saisi le montant** pendant le pointage
2. Vérifiez que le montant est un **nombre valide** (ex: 3000, pas "3000 FCFA")

### Problème: Après clôture, je ne vois plus les données
**Solution:**
1. Les données ne sont **pas supprimées**, elles sont dans la feuille "Archives"
2. Allez dans "Archives" pour les consulter
3. Les pointages sont juste **archivés**, pas perdus

---

## 📊 Exemples d'Utilisation Complète

### Scénario 1: Première réunion

**Étape 1**: Ajouter une réunion dans "Réunions"
```
Date: 15/09/2026
Lieu: Maison de Karim
Total collectés: 9000
Numéro: 1
```

**Étape 2**: Cliquer sur "Prendre les présences" → Sélectionner réunion 1

**Étape 3**: Pointer les membres
```
- KEITA Karim: Oui, 3000 FCFA
- DIALLO Mamadi: Oui, 3000 FCFA
- SOW Aïssatou: Non, 0 FCFA
```

**Résultat**: 
- Caisse = 6000 FCFA
- Pointages = 3 lignes ajoutées

---

### Scénario 2: Consulter le rapport d'un membre après 5 réunions

**Cliquer sur**: "Rapport Membre" → KEITA Karim

**Affichage**:
```
Membre: KEITA Karim

Nombre de réunion: 5/5
Nombre d'absences: 0

Total versé à la caisse: 15 000 FCFA
```

---

### Scénario 3: Clôturer l'année après 83 réunions

**Cliquer sur**: "Clôture l'année" → Confirmer

**Résultat**:
- Les 83 × 3 membres (environ 249 pointages) sont archivés
- Pointages est vide (prêt pour nouvelle année)
- Caisse continue (montant conservé)
- Réunions continue (pour l'historique)

---

## 💾 Sauvegarde

**Sauvegardez votre fichier** régulièrement:
1. `Ctrl + S` (Sauvegarde rapide)
2. `Ctrl + Maj + S` (Enregistrer sous)

**Format recommandé**: `.xlsm` (Excel avec macros)

---

## 📞 Support

Pour des questions ou des améliorations:
- Consultez le [GitHub](https://github.com/keitendul-glitch/gestion-reunion-vba)
- Ouvrez une issue avec vos questions

---

**Dernière mise à jour:** 2026-09-12

Bonne utilisation! 🎉
