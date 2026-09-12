# Gestion Réunion VBA

Un système complet de gestion des réunions en **VBA Excel** pour suivre les membres, pointages, dépenses et caisse.

## 📋 Fonctionnalités

✅ **Tableau de bord** - Affiche les données clés:
- Nombre de membres actifs
- Total en caisse
- Dernier lieu de réunion
- Boutons d'actions

✅ **Prendre les présences** - Boîte de dialogue pour:
- Sélectionner une réunion existante
- Ajouter les membres présents/absents
- Enregistrer les montants payés (3000 FCFA par défaut, ou acomptes 1000/2000 FCFA)

✅ **Rapport Membre** - Affiche pour chaque membre:
- Nombre total de réunions auxquelles il a participé
- Nombre total d'absences
- Total versé à la caisse
- Format: "66/83" (66 présences sur 83 réunions)

✅ **Clôture de l'année** - Archivage automatique:
- Déplace tous les pointages vers Archives
- Garde la liste des réunions intacte
- Réinitialise les pointages

## 📊 Structure des Feuilles

### Feuil1: Tableau de Bord
- Affichage des données clés
- Boutons d'actions (Présences, Rapport, Clôture)

### Feuil2: Membres
| Colonne | Données |
|---------|---------|
| A | Numéro membre |
| B | Nom |
| C | Prénom |
| D | Téléphone |
| E | Date inscription |

### Feuil3: Réunions
| Colonne | Données |
|---------|---------|
| A | Date |
| B | Lieu (hôte) |
| C | Total collectés |
| D | Numéro |

### Feuil4: Pointages
| Colonne | Données |
|---------|---------|
| A | Date |
| B | Nom et prénom |
| C | Présences (Oui/Non) |
| D | Montant versé |
| E | Lieu |

### Feuil5: Dépenses
| Colonne | Données |
|---------|---------|
| A | Date |
| B | Motif |
| C | Montant |

### Feuil6: Caisse
| Colonne | Données |
|---------|---------|
| A | Total caisse |

### Feuil7: Archives
| Colonne | Données |
|---------|---------|
| A | Date |
| B | Nom et prénom |
| C | Présences (Oui/Non) |
| D | Montant versé |
| E | Lieu |

## 🚀 Installation

1. Téléchargez le fichier `Gestion_Reunion.xlsm`
2. Ouvrez-le dans Excel
3. Les boutons du tableau de bord sont prêts à l'emploi
4. Activez les macros si demandé

## 💡 Utilisation

### Prendre les présences
1. Cliquez sur le bouton "Prendre les présences" du tableau de bord
2. Sélectionnez la réunion
3. Entrez chaque membre et son montant
4. Les données s'enregistrent automatiquement

### Consulter un rapport membre
1. Cliquez sur "Rapport Membre"
2. Entrez le nom et prénom du membre
3. Affichez ses statistiques complètes

### Clôturer l'année
1. Cliquez sur "Clôture l'année"
2. Confirmez l'archivage
3. Tous les pointages sont transférés aux Archives

## 📝 Notes

- Montant par défaut: **3000 FCFA**
- Acomptes acceptés: **1000 FCFA** ou **2000 FCFA**
- Format des présences: "Oui" ou "Non"
- Les données archivées peuvent être consultées en tout temps

## 👨‍💻 Auteur

Créé par keitendul-glitch

---

**Dernière mise à jour:** 2026-09-12
