VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SerpentForm 
   Caption         =   "Jeu du serpent"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9255.001
   OleObjectBlob   =   "SerpentForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SerpentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Labels :
'1 => tête
'2-129 => corps
'130-145 => pommes


Private Sub Label_fermer_Click()
    Unload Me
End Sub

Private Sub Label_fermer2_Click()
    Unload Me
End Sub



Private Sub Label_fond_Click()

End Sub

Private Sub Label_new_Click()
    
    Label_score.Caption = 0
    nb_pauses_partie = 0
    niveau = 1
    Label_niveau.Caption = niveau & "/5"
    
    reset_niveau
    
End Sub

Private Sub Label_new2_Click()
    Label_new_Click
End Sub

Private Sub reset_niveau()

    'Remettre en dehors
    For i = 2 To LastApple
        Controls("Label" & i).Visible = False
        Controls("Label" & i).Left = 1000
        Controls("Label" & i).Top = 1000
        If i <= SnakeHeadStop Then Controls("Label" & i).BackColor = &HC000&
    Next
    
    'Bleu en dehors
    Label_bleu.Left = 1000
    Label_bleu.Top = 1000
    
    Label1.BackColor = &HA000&
    Label1.Left = 7.5
    Label1.Top = 7.5
    
    UserForm_Initialize
    
End Sub

Private Sub Label_pause_Click()
    
    timer_debut_pause = 1
    
End Sub

Private Sub Label_pause2_Click()
    Label_pause_Click
End Sub

Private Sub Label_pause3_Click()
    Label_pause_Click
End Sub



Private Sub UserForm_Initialize()
    
    Randomize 'on initialise un générateur de nombre aléatoire
    
    

    fin_partie = False
    dernier_keycode = Empty
    timer_debut_partie = Empty 'Timer début de niveau
    jour_debut_partie = Empty 'Jour début de niveau
    timer_debut_pause = Empty
    nb_pommes_mangees = 0
    If niveau = Empty Then niveau = 1
    If nb_pauses_partie = Empty Then nb_pauses_partie = 0
    
    
    
    
    vitesse_initiale = 0.5
    nb_segments_serpent = 3 'Total : 128 + 1
    nb_pommes_en_jeu = 10
    For i = 1 To niveau - 1
        vitesse_initiale = vitesse_initiale - 0.05
        nb_segments_serpents = nb_segments_serpent + 2
        nb_pommes_en_jeu = nb_pommes_en_jeu - 2
    Next
    
        
    
    vitesse = vitesse_initiale
    
    'Serpent initial
    For i = 2 To nb_segments_serpent
        Controls("Label" & i).Visible = True
    Next
    
    'Tab
    For i = 1 To SnakeHeadStop
        tab_pos_x(i) = Controls("Label" & i).Left
        tab_pos_y(i) = Controls("Label" & i).Top
    Next
    
    
       
    
    'Pommes
    For pomme = 1 To nb_pommes_en_jeu
        
        erreur = False
        x_pomme = Int(33 * Rnd) * 9 + 7.5
        y_pomme = Int(33 * Rnd) * 9 + 7.5
    
        For i = 1 To FirstApple + pomme - 1
            If Controls("Label" & i).Visible = True Then
                If Controls("Label" & i).Left = x_pomme And Controls("Label" & i).Top = y_pomme Then
                    erreur = True
                End If
            End If
        Next
        
        If erreur Then 'Si déjà un label à cet emplacement
            pomme = pomme - 1
        Else 'Si OK
            Controls("Label" & pomme + SnakeHeadStop).Left = x_pomme
            Controls("Label" & pomme + SnakeHeadStop).Top = y_pomme
            Controls("Label" & pomme + SnakeHeadStop).Visible = True
        End If
        
    Next
    
End Sub

Private Sub replacer_pomme(ByVal p As Integer)

    For pomme = p To p
        
        erreur = False
        x_pomme = Int(33 * Rnd) * 9 + 7.5
        y_pomme = Int(33 * Rnd) * 9 + 7.5
    
        For i = 1 To LastApple
            If Controls("Label" & i).Visible = True Then
                If Controls("Label" & i).Left = x_pomme And Controls("Label" & i).Top = y_pomme Then
                    erreur = True
                End If
            End If
        Next
        
        If erreur Then 'Si déjà un label à cet emplacement
            pomme = pomme - 1
        Else 'Si OK
            Controls("Label" & pomme + LastApple).Left = x_pomme
            Controls("Label" & pomme + LastApple).Top = y_pomme
            Controls("Label" & pomme + LastApple).Visible = True
        End If
        
    Next
    
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Then
        
        'Test anti haut-bas droite-gauche
        annuler = False
        If Not dernier_keycode = Empty Then
            If (dernier_keycode = 37 And KeyCode = 39) Or (dernier_keycode = 38 And KeyCode = 40) Or (dernier_keycode = 39 And KeyCode = 37) Or (dernier_keycode = 40 And KeyCode = 38) Then
                annuler = True
            End If
        Else 'Si premier déplacement
            If KeyCode = 37 Or KeyCode = 38 Then
                annuler = True
            End If
        End If
        
        If annuler = False Then
            dernier_keycode = KeyCode
            prochain_depl_manuel = Timer
        End If
        
        'Fin de pause
        If Not timer_debut_pause = Empty Then
            If jour_debut_partie <> Day(Date) Then 'Si passage au jour suivant
                timer_debut_pause = timer_debut_pause - 86400
            End If
            timer_debut_partie = timer_debut_partie + (Timer - timer_debut_pause)
            timer_debut_pause = Empty
            demarrer
        End If
    
        'Start
        If timer_debut_partie = Empty Then
            timer_debut_partie = Timer
            jour_debut_partie = Day(Date)
            prochain_depl_auto = Timer + vitesse
            demarrer
        End If
        
    End If
    
    If KeyCode = 32 Then Label_pause_Click 'Barre d'espace => pause
    
End Sub

Private Sub test_jour()

    If jour_debut_partie <> Day(Date) Then 'Si passage au jour suivant
        timer_debut_partie = timer_debut_partie - 86400
        If Not prochain_depl_auto = Empty Then prochain_depl_auto = prochain_depl_auto - 86400
        jour_debut_partie = Day(Date)
    End If
    
End Sub

Sub demarrer()

    Do While fin_partie = False
        
        test_jour 'Anti bug passage d'un jour à l'autre
        
        deplacement = False
        
        If prochain_depl_auto <= Timer Then
            prochain_depl_auto = Timer + vitesse
            deplacement = True
        End If

        If prochain_depl_manuel <= Timer Then
            prochain_depl_manuel = 999999
            deplacement = True
            prochain_depl_auto = Timer + vitesse
        End If
        
        If timer_debut_pause <= 1 And timer_debut_pause <> 0 And nb_pauses_partie < 10 Then 'Si pause
        
            nb_pauses_partie = nb_pauses_partie + 1
            timer_debut_pause = Timer
            If nb_pauses_partie = 10 Then MsgBox "C'est votre 10e et dernière pause autorisée pour cette partie ;-)", 64, "Pause"
            Exit Do

        End If
        
        If deplacement Then
            
            On Error Resume Next
            
            'Vitesse
            vitesse = vitesse_initiale - Round(Sqr(Timer - timer_debut_partie) * (vitesse_initiale * 8), 0) / 100
            If vitesse < 0.05 Then vitesse = 0.05
            
            'Grille de 33x33
            decalage = 9
            
            depl_corps = False
            
            'Dépl. tête
            If dernier_keycode = 37 Then 'Gauche
                If Label1.Left > 7.5 Then
                    Label1.Left = Label1.Left - decalage
                    depl_corps = True
                Else
                    fin_partie = True
                End If
            ElseIf dernier_keycode = 38 Then 'Haut
                If Label1.Top > 7.5 Then
                    Label1.Top = Label1.Top - decalage
                    depl_corps = True
                Else
                    fin_partie = True
                End If
            ElseIf dernier_keycode = 39 Then 'Droite
                If Label1.Left < 292.5 Then
                    Label1.Left = Label1.Left + decalage
                    depl_corps = True
                Else
                    fin_partie = True
                End If
            ElseIf dernier_keycode = 40 Then 'Bas
                If Label1.Top < 292.5 Then
                    Label1.Top = Label1.Top + decalage
                    depl_corps = True
                Else
                    fin_partie = True
                End If
            End If
            
            'Dépl. corps
            If depl_corps Then
                
               
                
                'Dépl. du serpent (tête non comprise) + visibilité + test collision serpent
                For i = 2 To SnakeHeadStop
                    Controls("Label" & i).Left = tab_pos_x(i - 1)
                    Controls("Label" & i).Top = tab_pos_y(i - 1)
                    If Controls("Label" & i).Visible = True Then
                        If Controls("Label" & i).Left = Label1.Left And Controls("Label" & i).Top = Label1.Top Then
                            fin_partie = True
                            Controls("Label" & i).BackColor = &HDD&
                        End If
                    End If
                Next
                
                'Actu tab
                For i = 1 To SnakeHeadStop
                    tab_pos_x(i) = Controls("Label" & i).Left
                    tab_pos_y(i) = Controls("Label" & i).Top
                Next
                
                'Test collision pommes
                collision_pommes = False
                For i = FirstApple To LastApple
                    If Controls("Label" & i).Visible = True Then
                        If Controls("Label" & i).Left = Label1.Left And Controls("Label" & i).Top = Label1.Top Then 'Si pomme mangée
                        
                            Controls("Label" & i).Visible = False
                            If vitesse > vitesse_initiale * 0.4 Or niveau = 5 Then replacer_pomme i - LastApple
                            Label_score.Caption = Val(Label_score.Caption) + niveau * 5
                            nb_pommes_mangees = nb_pommes_mangees + 1.5
                            For ii = 2 To Round(nb_pommes_mangees, 0) + nb_segments_serpent
                                Controls("Label" & ii).Visible = True
                            Next
                            collision_pommes = True
                             
                        End If
                    End If
                Next

               'Test collision pomme bleue
                If Label_bleu.Visible = True Then
                    If Label_bleu.Left = Label1.Left And Label_bleu.Top = Label1.Top Then 'Si pomme mangée
                    
                        Label_bleu.Left = 1000
                        Label_bleu.Top = 1000
                        
                        Label_score.Caption = Val(Label_score.Caption) + niveau * 25
                        nb_pommes_mangees = nb_pommes_mangees + 1.5
                        For ii = 2 To Round(nb_pommes_mangees, 0) + nb_segments_serpent
                            Controls("Label" & ii).Visible = True
                        Next
                         
                    End If
                End If
                
                'Afficher/retirer pomme bleue
                nb_secondes_ecoulees = Timer - timer_debut_partie
                If vitesse > vitesse_initiale * 0.4 Or niveau = 5 Then
                    
                    If nb_secondes_ecoulees Mod 10 >= 5 Then 'Affichge
                    
                        If Label_bleu.Top < 1000 And Label_bleu.Visible = False Then 'Replacer
                            For pomme = 1 To 1
                                erreur = False
                                x_pomme = Int(33 * Rnd) * 9 + 7.5
                                y_pomme = Int(33 * Rnd) * 9 + 7.5
                            
                                For i = 1 To LastApple
                                    If Controls("Label" & i).Visible = True Then
                                        If Controls("Label" & i).Left = x_pomme And Controls("Label" & i).Top = y_pomme Then
                                            erreur = True
                                        End If
                                    End If
                                Next
                                
                                If erreur Then 'Si déjà un label à cet emplacement
                                    pomme = pomme - 1
                                Else 'Si OK
                                    Label_bleu.Left = x_pomme
                                    Label_bleu.Top = y_pomme
                                    Label_bleu.Visible = True
                                End If
                            Next
                        End If
                        
                    Else 'Masquer
                        Label_bleu.Visible = False
                        Label_bleu.Left = 900
                        Label_bleu.Top = 900
                    End If
                    
                End If

                'Test s'il reste des pommes
                nb_pommes_visibles = 0
                If collision_pommes Then
                    For i = FirstApple To LastApple
                        If Controls("Label" & i).Visible = True Then
                            nb_pommes_visibles = nb_pommes_visibles + 1
                        End If
                    Next
                    If nb_pommes_visibles = 0 Then 'Si plus aucune pomme
                        niveau = niveau + 1
                        Label_niveau.Caption = niveau & "/5"
                        reset_niveau
                        Exit Do
                    End If
                End If
                
            End If
        
            'Fin de partie
            If fin_partie Then
                
                duree_niveau_5 = Timer - timer_debut_partie
                
                Label1.BackColor = &HDD&
                
                Application.Wait Now + TimeValue("0:00:01")
                
                score = Val(Label_score.Caption)
               
                
                If score > score_10 And score > 0 Then
                    
                    pseudo = MsgBox("Félicitations, vous terminez avec un score de " & Label_score.Caption & " ! " & Chr(10) & Chr(10))
                    
                   
                    
                Else
                    If MsgBox("Vous terminez avec un score de " & Label_score.Caption & " ! " & Chr(10) & Chr(10) & "Ce score n'est pas suffisant pour être enregistré dans le top 10 ..." & Chr(10) & Chr(10) & "Souhaitez-vous rejouer ?", 36, "Fin de partie") = vbYes Then
                        Label_new_Click
                    End If
                End If
                
                niveau = 1
                Exit Do
                
            End If
        
        End If
        
        DoEvents
    
    Loop

End Sub



