# 🎯 Gestion Réunion VBA

> **Système complet de gestion des réunions en VBA Excel** - Suivi des membres, pointages, dépenses et caisse

[![GitHub](https://img.shields.io/badge/GitHub-View%20on%20GitHub-blue?logo=github)](https://github.com/keitendul-glitch/gestion-reunion-vba)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![Version](https://img.shields.io/badge/Version-1.0-blue)](.)

---

## 📋 Fonctionnalités Principales

### ✅ Tableau de Bord
- 📊 Nombre de membres actifs
- 💰 Total en caisse en temps réel
- 📍 Dernier lieu de réunion
- 🔘 3 boutons d'actions

### ✅ Prendre les Présences
- 📝 Boîte de dialogue interactive
- ➕ Ajouter présences/absences par réunion
- 💵 Enregistrement des montants versés
- 🔄 Mise à jour automatique de la caisse

### ✅ Rapport Membre
- 📈 Nombre de réunions participées
- ❌ Nombre d'absences
- 💸 Total versé à la caisse
- Format: **66/83** (66 participations sur 83 réunions)

### ✅ Clôture de l'Année
- 📦 Archivage automatique des pointages
- 🔒 Conservation de l'historique
- 🔄 Réinitialisation pour nouvelle année
- 📊 Réunions et caisse intactes

---

## 🚀 Démarrage Rapide

### Option 1️⃣: Installation Manuelle (5 minutes)

1. **Créez un fichier Excel** avec 7 feuilles:
   - Tableau de bord
   - Membres
   - Réunions
   - Pointages
   - Dépenses
   - Caisse
   - Archives

2. **Importez le code VBA**:
   - Ouvrez l'Éditeur VBA (`Alt + F11`)
   - Créez un nouveau module
   - Copiez le code de `VBA/Module_Principal.vb`

3. **Créez 3 boutons** sur le tableau de bord:
   - Prendre les présences → `BoutonPrendrePresences`
   - Rapport Membre → `BoutonRapportMembre`
   - Clôture l'année → `BoutonClotureAnnee`

👉 **Guide détaillé**: [GUIDE_INSTALLATION.md](docs/GUIDE_INSTALLATION.md)

### Option 2️⃣: Utiliser les Données d'Exemple

Consultez [DONNEES_EXEMPLE.md](docs/DONNEES_EXEMPLE.md) pour voir:
- ✅ La structure complète du classeur
- ✅ Des exemples de données réelles
- ✅ Des rapports générés automatiquement

---

## 📖 Documentation

| Document | Description |
|----------|-------------|
| [GUIDE_INSTALLATION.md](docs/GUIDE_INSTALLATION.md) | 🚀 Installation étape par étape |
| [CONFIGURATION.md](docs/CONFIGURATION.md) | ⚙️ Paramètres et personnalisation |
| [DONNEES_EXEMPLE.md](docs/DONNEES_EXEMPLE.md) | 📊 Données d'exemple et structure |
| [DEPANNAGE.md](docs/DEPANNAGE.md) | 🆘 Guide de dépannage complet |

---

## 💡 Utilisation

### 1. Prendre les Présences

```
1. Cliquez sur "Prendre les présences"
2. Sélectionnez une réunion existante
3. Pour chaque membre:
   - Entrez "Oui" ou "Non" pour la présence
   - Entrez le montant (3000 FCFA par défaut)
4. Les données s'enregistrent automatiquement
5. La caisse se met à jour
```

**Montants acceptés:**
- 3000 FCFA (montant standard)
- 1000 FCFA (acompte)
- 2000 FCFA (acompte)
- Tout autre montant personnalisé

### 2. Consulter un Rapport

```
1. Cliquez sur "Rapport Membre"
2. Entrez le NOM: KEITA
3. Entrez le PRÉNOM: Karim
4. Résultat:
   Nombre de réunion: 66/83
   Nombre d'absences: 17
   Total versé à la caisse: 198 000 FCFA
```

### 3. Clôturer l'Année

```
1. Cliquez sur "Clôture l'année"
2. Confirmez l'archivage
3. Tous les pointages → Archives
4. Feuille Pointages → vide (prête pour nouvelle année)
5. Réunions et Caisse → conservées
```

---

## 📊 Structure des Données

### Feuil "Membres"
```
Numéro membre | Nom | Prénom | Téléphone | Date inscription
```

### Feuil "Réunions"
```
Date | Lieu | Total collectés | Numéro
```

### Feuil "Pointages"
```
Date | Nom et prénom | Présences | Montant versé | Lieu
```

### Feuil "Dépenses"
```
Date | Motif | Montant
```

### Feuil "Caisse"
```
Total caisse
```

### Feuil "Archives"
```
Date | Nom et prénom | Présences | Montant versé | Lieu
```

---

## 🔧 Fichiers du Projet

```
gestion-reunion-vba/
├── README.md                          # Ce fichier
├── VBA/
│   └── Module_Principal.vb            # Code VBA complet
└── docs/
    ├── GUIDE_INSTALLATION.md          # Installation détaillée
    ├── CONFIGURATION.md               # Paramètres
    ├── DONNEES_EXEMPLE.md             # Données d'exemple
    └── DEPANNAGE.md                   # Guide de dépannage
```

---

## 🎯 Exemples

### Exemple 1: Ajouter un nouveau pointage
```
Réunion: 1. 15/09/2026 - Maison de Karim

Membre: KEITA Karim
  Présent? Oui
  Montant? 3000
  
Membre: DIALLO Mamadi
  Présent? Oui
  Montant? 2000 (acompte)
  
Membre: SOW Aïssatou
  Présent? Non
  Montant? 0
```

### Exemple 2: Consulter un rapport
```
Membre: KEITA Karim

Résultat:
  Nombre de réunion: 66/83
  Nombre d'absences: 17
  Total versé à la caisse: 198 000 FCFA
```

### Exemple 3: Clôturer l'année
```
Avant:
  - Pointages: 249 lignes (83 réunions × 3 membres)
  - Caisse: 249 000 FCFA

Après clôture:
  - Pointages: vide (prêt pour nouvelle année)
  - Archives: 249 lignes (historique conservé)
  - Caisse: 249 000 FCFA (conservé)
  - Réunions: conservé
```

---

## ⚙️ Configuration Personnalisée

### Modifier le montant par défaut
Fichier: `VBA/Module_Principal.vb` (ligne ~220)
```vb
If montantInput = "" Then
    montant = 3000  ' ← MODIFIER CETTE VALEUR
```

### Modifier les noms des feuilles
Assurez-vous que les noms correspondent exactement au code VBA:
- Tableau de bord
- Membres
- Réunions
- Pointages
- Dépenses
- Caisse
- Archives

---

## 🐛 Dépannage

### Les boutons ne fonctionnent pas?
→ Activez les macros: `Fichier` → `Options` → `Centre de gestion de la confidentialité` → `Activer les macros`

### Le rapport affiche 0?
→ Vérifiez que le nom et prénom correspondent exactement aux données

### La caisse n'augmente pas?
→ Assurez-vous d'avoir saisi un montant lors du pointage

👉 **Guide complet**: [DEPANNAGE.md](docs/DEPANNAGE.md)

---

## 📞 Support

### Documentation
- 📖 Lisez le [GUIDE_INSTALLATION.md](docs/GUIDE_INSTALLATION.md)
- 🔍 Consultez le [DEPANNAGE.md](docs/DEPANNAGE.md)
- 📊 Voir les [DONNEES_EXEMPLE.md](docs/DONNEES_EXEMPLE.md)

### GitHub
- 🐛 [Ouvrir une Issue](https://github.com/keitendul-glitch/gestion-reunion-vba/issues)
- ⭐ [Mettre en favori le projet](https://github.com/keitendul-glitch/gestion-reunion-vba)

---

## 📝 Notes

- ✅ **Montant par défaut**: 3000 FCFA
- ✅ **Format dates**: DD/MM/YYYY
- ✅ **Format présences**: "Oui" ou "Non"
- ✅ **Sauvegarde**: Format `.xlsm` (Excel avec macros)
- ✅ **Données archivées**: Conservées après clôture

---

## 🎓 Versions

| Version | Date | Changements |
|---------|------|------------|
| 1.0 | 2026-09-12 | 🎉 Première version complète |

---

## 👨‍💻 Auteur

Créé par **keitendul-glitch**

---

## 📜 Licence

Ce projet est sous licence **MIT**. Vous êtes libre de l'utiliser, le modifier et le distribuer.

---

## 🎉 Bon Courage!

**Merci d'utiliser Gestion Réunion VBA!**

Pour commencer → [GUIDE_INSTALLATION.md](docs/GUIDE_INSTALLATION.md)

---

**Dernière mise à jour:** 2026-09-12
