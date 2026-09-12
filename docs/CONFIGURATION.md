# Configuration et Paramètres

## ⚙️ Paramètres Modifiables

### Montants par défaut

Si vous souhaitez **modifier le montant par défaut** (actuellement 3000 FCFA):

**Fichier**: `VBA/Module_Principal.vb`
**Ligne**: Environ 220

```vb
If montantInput = "" Then
    montant = 3000  ' ← MODIFIER CETTE VALEUR
Else
```

**Exemples**:
- 2500 FCFA: `montant = 2500`
- 5000 FCFA: `montant = 5000`

---

## 📝 Variables Système

### Noms des feuilles

Si vous renommez les feuilles, **mettez à jour le code VBA**:

```vb
Set ws = ThisWorkbook.Sheets("Tableau de bord")  ' Feuil1
Set wsReunions = ThisWorkbook.Sheets("Réunions")  ' Feuil3
Set wsMembres = ThisWorkbook.Sheets("Membres")    ' Feuil2
Set wsPointages = ThisWorkbook.Sheets("Pointages")' Feuil4
Set wsCaisse = ThisWorkbook.Sheets("Caisse")      ' Feuil6
Set wsArchives = ThisWorkbook.Sheets("Archives")  ' Feuil7
```

---

## 🔐 Sécurité

### Protéger les feuilles importantes

1. Cliquez sur la feuille "Caisse"
2. Onglet "Révision" → "Protéger la feuille"
3. Entrez un mot de passe (optionnel)

**Feuilles recommandées à protéger:**
- Caisse
- Archives

---

## 📊 Formules Avancées

### Ajouter un total automatique dans "Caisse"

1. Allez à la feuille "Pointages"
2. Créez une formule dans une cellule libre:

```excel
=SOMME(D:D)
```

Cette formule additionne tous les montants versés.

---

## 🐛 Débogage

### Afficher les erreurs dans le code VBA

Si une macro plante, le message d'erreur apparaît. **Notez le message** et consultez:

**Erreurs courantes:**

| Erreur | Cause | Solution |
|--------|-------|----------|
| "Feuille non trouvée" | Nom de feuille incorrecte | Vérifiez les noms des feuilles |
| "Indice en dehors des limites" | Cellule vide ou manquante | Ajoutez des données d'exemple |
| "Chaîne de caractères non valide" | Format de date incorrect | Utilisez DD/MM/YYYY |

---

## ♻️ Réinitialiser le système

### Si vous voulez recommencer de zéro:

1. **Videz toutes les feuilles** sauf les en-têtes:
   - Pointages (garder ligne 1)
   - Réunions (garder ligne 1)
   - Dépenses (garder ligne 1)

2. **Réinitialisez la caisse**:
   - Feuille "Caisse" → A2 → Effacez le contenu

3. **Gardez les membres intacts** (Feuil2)

4. **Les archives restent** pour l'historique

---

## 📈 Améliorations Possibles

### Fonctionnalités futures à ajouter:

- ✏️ Modifier les pointages existants
- 🔍 Rechercher un membre par critère
- 📊 Générer des rapports visuels (graphiques)
- 📧 Exporter les données en PDF
- 🔔 Alertes pour les absences répétées
- 💾 Sauvegarde automatique

---

## 🔄 Mise à jour du code

### Comment mettre à jour le code VBA:

1. Ouvrez l'**Éditeur VBA** (`Alt + F11`)
2. Sélectionnez le module `Module_Principal`
3. **Supprimez tout** le code existant
4. **Collez le nouveau code** du GitHub
5. **Sauvegardez** (`Ctrl + S`)

---

**Dernière mise à jour:** 2026-09-12
