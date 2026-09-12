# 🆘 Guide de Dépannage

## ❌ Problèmes Courants et Solutions

---

## 1. Les macros ne s'exécutent pas

### Symptôme
Les boutons ne réagissent pas au clic.

### Causes possibles
- Les macros sont **désactivées**
- Le fichier n'est pas en format `.xlsm`

### Solutions

**Solution A: Activer les macros**
1. Ouvrez le fichier Excel
2. Si une barre jaune apparaît en haut → Cliquez sur "Activer le contenu"
3. Acceptez l'activation des macros

**Solution B: Changer le format du fichier**
1. `Fichier` → `Enregistrer sous`
2. Format: Sélectionnez **"Classeur Excel prenant en charge les macros (.xlsm)"`
3. Cliquez sur "Enregistrer"

**Solution C: Activer les macros dans les paramètres**
1. `Fichier` → `Options` → `Centre de gestion de la confidentialité`
2. Cliquez sur "Paramètres du Centre de gestion de la confidentialité"
3. Onglet "Paramètres des macros"
4. Cochez **"Activer toutes les macros"** (ou au minimum "Afficher une notification")
5. Cliquez sur "OK"

---

## 2. Erreur: "Feuille non trouvée"

### Symptôme
Message d'erreur: `"Erreur: Feuille non trouvée"`

### Cause
Les noms des feuilles ne correspondent pas au code.

### Solution
1. Ouvrez l'**Éditeur VBA** (`Alt + F11`)
2. Allez dans le fichier `VBA/Module_Principal.vb` dans le GitHub
3. Vérifiez les noms de feuilles au début du code
4. Renommez vos feuilles Excel exactement comme dans le code:
   - `Tableau de bord`
   - `Membres`
   - `Réunions`
   - `Pointages`
   - `Dépenses`
   - `Caisse`
   - `Archives`

---

## 3. Le rapport affiche "0" pour tout

### Symptôme
Quand vous consultez le rapport d'un membre, tout affiche 0.

### Causes possibles
- Le **nom et prénom ne correspondent pas exactement**
- Les données ne sont **pas dans la bonne feuille**
- La feuille "Pointages" est vide

### Solutions

**Solution A: Vérifier l'orthographe**
1. Allez à la feuille "Pointages"
2. Vérifiez la colonne B (Nom et prénom)
3. Le rapport utilise "Nom" + "Prénom" (séparés par un espace)
4. Exemple: Si la feuille affiche "KEITA Karim", entrez exactement la même chose

**Solution B: Vérifier les données**
1. Allez à la feuille "Pointages"
2. Vérifiez qu'il y a au moins une ligne de données (à partir de la ligne 2)
3. Si vide, utilisez "Prendre les présences" pour ajouter des données

**Solution C: Vérifier le format des montants**
1. Colonne D (Montant versé) doit contenir des **nombres**
2. Si c'est du texte (ex: "3000 FCFA"), changez-le en nombre: `3000`

---

## 4. La caisse n'augmente pas après pointage

### Symptôme
Après avoir pris les présences et entré les montants, le total en caisse ne change pas.

### Causes possibles
- Vous n'avez pas **saisi le montant** lors du pointage
- Le montant est en format **texte** au lieu de nombre

### Solutions

**Solution A: Refaire le pointage avec montants**
1. Cliquez sur "Prendre les présences"
2. À chaque membre, assurez-vous d'entrer un montant
3. Ne laissez pas le champ vide

**Solution B: Vérifier la feuille Caisse**
1. Allez à la feuille "Caisse"
2. Cellule A2 devrait contenir le total
3. Si vide ou 0, c'est que aucun montant n'a été enregistré

**Solution C: Vérifier le format des montants**
1. Allez à "Pointages" → Colonne D
2. Sélectionnez une cellule avec un montant
3. Clic droit → "Format de cellule"
4. Format: Doit être **"Nombre"** (pas "Texte")

---

## 5. Erreur lors de la clôture de l'année

### Symptôme
Message d'erreur pendant "Clôture l'année"

### Cause
Probablement un problème de format de données dans la feuille "Pointages"

### Solutions

1. Avant de clôturer, vérifiez que:
   - Colonne A (Date): Format date valide
   - Colonne B (Nom et prénom): Texte
   - Colonne C (Présences): "Oui" ou "Non"
   - Colonne D (Montant): Nombre
   - Colonne E (Lieu): Texte

2. Si l'erreur persiste:
   - Allez à "Pointages" → Sélectionnez les lignes avec données
   - `Données` → `Valider` → Vérifiez qu'il n'y a pas de contraintes

---

## 6. Les en-têtes des feuilles ont disparu

### Symptôme
Quand vous ouvrez la feuille, la ligne 1 avec les noms des colonnes n'est plus là.

### Cause
Vous avez accidentellement supprimé la ligne 1.

### Solution

1. Récupérez à partir de la sauvegarde précédente:
   - `Fichier` → `Infos` → `Versions précédentes` (si disponible)

2. Ou recréez les en-têtes manuellement:
   - Sélectionnez la ligne 1
   - Clic droit → "Insérer au-dessus"
   - Recopiez les en-têtes

---

## 7. "Aucune réunion disponible"

### Symptôme
Quand vous cliquez sur "Prendre les présences", un message dit "Aucune réunion disponible"

### Cause
La feuille "Réunions" ne contient pas de données.

### Solution

1. Allez à la feuille "Réunions"
2. Ajoutez au moins une réunion:
   ```
   Ligne 1 (En-têtes): Date | Lieu | Total collectés | Numéro
   Ligne 2 (Données):  15/09/2026 | Maison de Karim | 9000 | 1
   ```
3. Réessayez "Prendre les présences"

---

## 8. "Aucun membre disponible"

### Symptôme
Quand vous prenez les présences, un message dit "Aucun membre disponible"

### Cause
La feuille "Membres" ne contient pas de données.

### Solution

1. Allez à la feuille "Membres"
2. Ajoutez au moins un membre:
   ```
   Ligne 1 (En-têtes): Numéro membre | Nom | Prénom | Téléphone | Date inscription
   Ligne 2 (Données): 1 | KEITA | Karim | 0611223344 | 15/01/2020
   ```
3. Réessayez "Prendre les présences"

---

## 9. L'année s'est clôturée par erreur

### Symptôme
Vous avez clôturé l'année accidentellement et les pointages ont disparu.

### Solution (Récupération)

**BONNE NOUVELLE**: Les données sont dans "Archives"!

1. Allez à la feuille "Archives"
2. Sélectionnez toutes les données archivées (Ctrl + A)
3. Copiez (Ctrl + C)
4. Allez à "Pointages"
5. Collez (Ctrl + V) à partir de la ligne 2
6. Les données sont récupérées!

**Pour éviter cela à l'avenir:**
- Toujours faire une **sauvegarde** avant de clôturer
- Utiliser `Ctrl + Z` immédiatement après l'erreur

---

## 10. Le fichier est très lent

### Symptôme
Excel ralentit ou fige quand vous utilisez les macros.

### Cause
Trop de données dans les feuilles (plusieurs années d'archives).

### Solutions

**Solution A: Nettoyer les Archives**
1. Allez à "Archives"
2. Supprimez les lignes les plus anciennes
3. Conservez seulement l'historique récent

**Solution B: Créer un nouveau fichier**
1. Créez un nouveau classeur pour la nouvelle année
2. Gardez l'ancien fichier comme archive

**Solution C: Optimiser Excel**
1. `Fichier` → `Options` → `Avancées`
2. Réduisez les options de recalcul
3. Désactivez les animations

---

## 📞 Besoin d'aide supplémentaire?

Si votre problème ne figure pas ici:

1. **Consultez le code VBA** sur GitHub:
   - https://github.com/keitendul-glitch/gestion-reunion-vba

2. **Vérifiez le format de vos données**:
   - Dates en DD/MM/YYYY
   - Montants en nombres (pas texte)
   - Noms exactement comme saisis

3. **Testez avec les données d'exemple** du guide

4. **Ouvrez une issue** sur GitHub avec:
   - Le message d'erreur exact
   - Qu'avez-vous fait quand l'erreur est survenue
   - Une capture d'écran si possible

---

**Dernière mise à jour:** 2026-09-12

Bon courage! 💪
