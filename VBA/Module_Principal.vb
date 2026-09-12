' ===================================================
' GESTION RÉUNION VBA - CODE PRINCIPAL
' ===================================================
' Système complet de gestion des réunions en Excel
' Auteur: keitendul-glitch
' Date: 2026-09-12
' ===================================================

Option Explicit

' ===================================================
' 1. INITIALISATION DU TABLEAU DE BORD
' ===================================================
Sub InitialiserTableauDeBord()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tableau de bord")
    
    ' Nettoyer le tableau de bord
    ws.Range("A1:E20").Clear
    
    ' Titres
    ws.Range("A1").Value = "TABLEAU DE BORD - GESTION RÉUNION"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    
    ' Afficher les données
    ws.Range("A3").Value = "Nombre de membres actifs:"
    ws.Range("B3").Value = CompterMembresActifs()
    
    ws.Range("A4").Value = "Total en caisse:"
    ws.Range("B4").Value = GetTotalCaisse()
    ws.Range("B4").NumberFormat = "0 ""FCFA"""
    
    ws.Range("A5").Value = "Dernier lieu de réunion:"
    ws.Range("B5").Value = GetDernierLieuReunion()
    
    ' Espacer les boutons
    ws.Range("A7").Value = "ACTIONS:"
    ws.Range("A7").Font.Bold = True
    
End Sub

' ===================================================
' 2. COMPTER LES MEMBRES ACTIFS
' ===================================================
Function CompterMembresActifs() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Membres")
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    CompterMembresActifs = lastRow - 1 ' Exclure l'en-tête
End Function

' ===================================================
' 3. OBTENIR LE TOTAL EN CAISSE
' ===================================================
Function GetTotalCaisse() As Double
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Caisse")
    
    If ws.Range("A2").Value = "" Then
        GetTotalCaisse = 0
    Else
        GetTotalCaisse = ws.Range("A2").Value
    End If
End Function

' ===================================================
' 4. OBTENIR LE DERNIER LIEU DE RÉUNION
' ===================================================
Function GetDernierLieuReunion() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Réunions")
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow = 1 Then
        GetDernierLieuReunion = "Aucune réunion"
    Else
        GetDernierLieuReunion = ws.Cells(lastRow, 2).Value ' Colonne B (Lieu)
    End If
End Function

' ===================================================
' 5. BOUTON: PRENDRE LES PRÉSENCES
' ===================================================
Sub BoutonPrendrePresences()
    Call AfficherFormulairePrecedences
End Sub

' ===================================================
' 6. FORMULAIRE PRENDRE LES PRÉSENCES
' ===================================================
Sub AfficherFormulairePrecedences()
    Dim wsReunions As Worksheet, wsMembres As Worksheet, wsPointages As Worksheet
    Dim lastRowReunions As Long, lastRowMembres As Long
    Dim i As Long
    Dim inputReunion As String
    Dim indexReunion As Long
    Dim dateReunion As String, lieuReunion As String
    
    Set wsReunions = ThisWorkbook.Sheets("Réunions")
    Set wsMembres = ThisWorkbook.Sheets("Membres")
    Set wsPointages = ThisWorkbook.Sheets("Pointages")
    
    lastRowReunions = wsReunions.Cells(wsReunions.Rows.Count, "A").End(xlUp).Row
    lastRowMembres = wsMembres.Cells(wsMembres.Rows.Count, "A").End(xlUp).Row
    
    If lastRowReunions = 1 Then
        MsgBox "Aucune réunion disponible. Créez d'abord une réunion dans la feuille 'Réunions'.", vbExclamation
        Exit Sub
    End If
    
    If lastRowMembres = 1 Then
        MsgBox "Aucun membre disponible. Ajoutez d'abord des membres dans la feuille 'Membres'.", vbExclamation
        Exit Sub
    End If
    
    ' Créer la liste des réunions
    Dim listeReunions As String
    listeReunions = ""
    For i = 2 To lastRowReunions
        listeReunions = listeReunions & (i - 1) & ". " & Format(wsReunions.Cells(i, 1).Value, "DD/MM/YYYY") & " - " & wsReunions.Cells(i, 2).Value & vbCrLf
    Next i
    
    ' Demander quelle réunion
    inputReunion = InputBox("Sélectionnez la réunion (numéro):" & vbCrLf & vbCrLf & listeReunions, "Prendre les présences")
    
    If inputReunion = "" Then Exit Sub
    
    On Error Resume Next
    indexReunion = CLng(inputReunion)
    On Error GoTo 0
    
    If indexReunion < 1 Or indexReunion > lastRowReunions - 1 Then
        MsgBox "Numéro de réunion invalide.", vbExclamation
        Exit Sub
    End If
    
    dateReunion = wsReunions.Cells(indexReunion + 1, 1).Value
    lieuReunion = wsReunions.Cells(indexReunion + 1, 2).Value
    
    ' Afficher les membres pour pointage
    Call AjouterPresences(dateReunion, lieuReunion, lastRowMembres)
    
End Sub

' ===================================================
' 7. AJOUTER LES PRÉSENCES
' ===================================================
Sub AjouterPresences(dateReunion As String, lieuReunion As String, lastRowMembres As Long)
    Dim wsPointages As Worksheet, wsMembres As Worksheet
    Dim i As Long, lastRowPointages As Long
    Dim nomMembre As String, presence As String
    Dim montantInput As String, montant As Double
    
    Set wsPointages = ThisWorkbook.Sheets("Pointages")
    Set wsMembres = ThisWorkbook.Sheets("Membres")
    
    lastRowPointages = wsPointages.Cells(wsPointages.Rows.Count, "A").End(xlUp).Row
    If lastRowPointages = 1 Then lastRowPointages = 1 ' Si vide, commencer à 1
    
    ' Boucle pour chaque membre
    For i = 2 To lastRowMembres
        nomMembre = wsMembres.Cells(i, 2).Value & " " & wsMembres.Cells(i, 3).Value
        
        ' Demander présence
        presence = InputBox("Est-ce que " & nomMembre & " était présent?" & vbCrLf & "(Tapez: Oui ou Non)" & vbCrLf & vbCrLf & "(Appuyez sur Annuler pour quitter)", "Pointage - " & nomMembre)
        
        If presence = "" Then Exit For
        
        If LCase(presence) = "oui" Or LCase(presence) = "non" Then
            ' Demander le montant versé
            montantInput = InputBox("Montant versé par " & nomMembre & " (FCFA):" & vbCrLf & "(Par défaut: 3000, acomptes: 1000 ou 2000)", "Montant")
            
            If montantInput = "" Then
                montant = 3000
            Else
                On Error Resume Next
                montant = CDbl(montantInput)
                On Error GoTo 0
                If montant = 0 Then montant = 3000
            End If
            
            ' Ajouter à Pointages
            lastRowPointages = lastRowPointages + 1
            wsPointages.Cells(lastRowPointages, 1).Value = dateReunion ' Date
            wsPointages.Cells(lastRowPointages, 2).Value = nomMembre ' Nom et prénom
            wsPointages.Cells(lastRowPointages, 3).Value = presence ' Présence
            wsPointages.Cells(lastRowPointages, 4).Value = montant ' Montant
            wsPointages.Cells(lastRowPointages, 5).Value = lieuReunion ' Lieu
            
            ' Mettre à jour la caisse
            Call AjouterALaCaisse(montant)
            
            MsgBox nomMembre & " enregistré(e) (" & presence & ") - " & montant & " FCFA", vbInformation
        Else
            MsgBox "Veuillez entrer 'Oui' ou 'Non'", vbExclamation
            i = i - 1 ' Refaire le pointage pour ce membre
        End If
    Next i
    
    MsgBox "Pointage terminé pour cette réunion!", vbInformation
    Call InitialiserTableauDeBord ' Actualiser le tableau de bord
End Sub

' ===================================================
' 8. AJOUTER À LA CAISSE
' ===================================================
Sub AjouterALaCaisse(montant As Double)
    Dim wsCaisse As Worksheet
    Set wsCaisse = ThisWorkbook.Sheets("Caisse")
    
    Dim totalActuel As Double
    totalActuel = GetTotalCaisse()
    wsCaisse.Range("A2").Value = totalActuel + montant
End Sub

' ===================================================
' 9. BOUTON: RAPPORT MEMBRE
' ===================================================
Sub BoutonRapportMembre()
    Call AfficherRapportMembre
End Sub

' ===================================================
' 10. AFFICHER LE RAPPORT MEMBRE
' ===================================================
Sub AfficherRapportMembre()
    Dim nom As String, prenom As String
    Dim nombreParticipations As Long, nombreAbsences As Long
    Dim totalVerse As Double
    Dim nombreTotalReunions As Long
    Dim nomComplet As String
    
    ' Demander le nom et prénom
    nom = InputBox("Entrez le NOM du membre:", "Rapport Membre")
    If nom = "" Then Exit Sub
    
    prenom = InputBox("Entrez le PRÉNOM du membre:", "Rapport Membre")
    If prenom = "" Then Exit Sub
    
    nomComplet = nom & " " & prenom
    
    ' Calculer les statistiques
    nombreParticipations = CompterParticipations(nomComplet)
    nombreAbsences = CompterAbsences(nomComplet)
    totalVerse = CalculerTotalVerse(nomComplet)
    nombreTotalReunions = CompterTotalReunions()
    
    ' Afficher le rapport
    Dim rapport As String
    rapport = "Membre: " & nomComplet & vbCrLf & vbCrLf
    rapport = rapport & "Nombre de réunion: " & nombreParticipations & "/" & nombreTotalReunions & vbCrLf
    rapport = rapport & "Nombre d'absences: " & nombreAbsences & vbCrLf & vbCrLf
    rapport = rapport & "Total versé à la caisse: " & Format(totalVerse, "0 ""FCFA""")
    
    MsgBox rapport, vbInformation, "Rapport - " & nomComplet
End Sub

' ===================================================
' 11. COMPTER LES PARTICIPATIONS
' ===================================================
Function CompterParticipations(nomComplet As String) As Long
    Dim wsPointages As Worksheet
    Dim lastRow As Long, i As Long
    Dim count As Long
    
    Set wsPointages = ThisWorkbook.Sheets("Pointages")
    lastRow = wsPointages.Cells(wsPointages.Rows.Count, "A").End(xlUp).Row
    
    count = 0
    For i = 2 To lastRow
        If wsPointages.Cells(i, 2).Value = nomComplet And LCase(wsPointages.Cells(i, 3).Value) = "oui" Then
            count = count + 1
        End If
    Next i
    
    CompterParticipations = count
End Function

' ===================================================
' 12. COMPTER LES ABSENCES
' ===================================================
Function CompterAbsences(nomComplet As String) As Long
    Dim wsPointages As Worksheet
    Dim lastRow As Long, i As Long
    Dim count As Long
    
    Set wsPointages = ThisWorkbook.Sheets("Pointages")
    lastRow = wsPointages.Cells(wsPointages.Rows.Count, "A").End(xlUp).Row
    
    count = 0
    For i = 2 To lastRow
        If wsPointages.Cells(i, 2).Value = nomComplet And LCase(wsPointages.Cells(i, 3).Value) = "non" Then
            count = count + 1
        End If
    Next i
    
    CompterAbsences = count
End Function

' ===================================================
' 13. CALCULER LE TOTAL VERSÉ
' ===================================================
Function CalculerTotalVerse(nomComplet As String) As Double
    Dim wsPointages As Worksheet
    Dim lastRow As Long, i As Long
    Dim total As Double
    
    Set wsPointages = ThisWorkbook.Sheets("Pointages")
    lastRow = wsPointages.Cells(wsPointages.Rows.Count, "A").End(xlUp).Row
    
    total = 0
    For i = 2 To lastRow
        If wsPointages.Cells(i, 2).Value = nomComplet Then
            total = total + wsPointages.Cells(i, 4).Value
        End If
    Next i
    
    CalculerTotalVerse = total
End Function

' ===================================================
' 14. COMPTER TOTAL RÉUNIONS
' ===================================================
Function CompterTotalReunions() As Long
    Dim wsReunions As Worksheet
    Dim lastRow As Long
    
    Set wsReunions = ThisWorkbook.Sheets("Réunions")
    lastRow = wsReunions.Cells(wsReunions.Rows.Count, "A").End(xlUp).Row
    
    CompterTotalReunions = IIf(lastRow > 1, lastRow - 1, 0)
End Function

' ===================================================
' 15. BOUTON: CLÔTURE DE L'ANNÉE
' ===================================================
Sub BoutonClotureAnnee()
    Call ClotureAnnee
End Sub

' ===================================================
' 16. CLÔTURE DE L'ANNÉE
' ===================================================
Sub ClotureAnnee()
    Dim response As Long
    Dim wsPointages As Worksheet, wsArchives As Worksheet
    Dim lastRowPointages As Long, lastRowArchives As Long
    Dim i As Long
    
    Set wsPointages = ThisWorkbook.Sheets("Pointages")
    Set wsArchives = ThisWorkbook.Sheets("Archives")
    
    ' Confirmation
    response = MsgBox("Êtes-vous sûr de vouloir clôturer l'année?" & vbCrLf & "Tous les pointages seront archivés.", vbYesNo, "Clôture de l'année")
    
    If response = vbNo Then Exit Sub
    
    lastRowPointages = wsPointages.Cells(wsPointages.Rows.Count, "A").End(xlUp).Row
    lastRowArchives = wsArchives.Cells(wsArchives.Rows.Count, "A").End(xlUp).Row
    
    If lastRowPointages <= 1 Then
        MsgBox "Aucun pointage à archiver.", vbInformation
        Exit Sub
    End If
    
    ' Copier tous les pointages vers Archives
    For i = 2 To lastRowPointages
        lastRowArchives = lastRowArchives + 1
        wsArchives.Cells(lastRowArchives, 1).Value = wsPointages.Cells(i, 1).Value ' Date
        wsArchives.Cells(lastRowArchives, 2).Value = wsPointages.Cells(i, 2).Value ' Nom et prénom
        wsArchives.Cells(lastRowArchives, 3).Value = wsPointages.Cells(i, 3).Value ' Présence
        wsArchives.Cells(lastRowArchives, 4).Value = wsPointages.Cells(i, 4).Value ' Montant
        wsArchives.Cells(lastRowArchives, 5).Value = wsPointages.Cells(i, 5).Value ' Lieu
    Next i
    
    ' Vider les pointages (garder l'en-tête)
    If lastRowPointages > 1 Then
        wsPointages.Rows("2:" & lastRowPointages).Delete
    End If
    
    MsgBox "Année clôturée!" & vbCrLf & (lastRowPointages - 1) & " enregistrements ont été archivés.", vbInformation
    Call InitialiserTableauDeBord
End Sub

' ===================================================
' FIN DU CODE
' ===================================================
