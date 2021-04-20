Attribute VB_Name = "Serpent"

Public Const SnakeHeadStart As Integer = 1, SnakeHeadStop As Integer = 129, FirstApple As Integer = 130, LastApple As Integer = 145

Public dernier_keycode, fin_partie, vitesse, vitesse_initiale, timer_debut_partie, jour_debut_partie, _
prochain_depl_manuel, prochain_depl_auto, nb_segments_serpent, tab_pos_x(SnakeHeadStop), tab_pos_y(SnakeHeadStop), _
nb_pommes_en_jeu, nb_pommes_mangees, niveau, timer_debut_pause, nb_pauses_partie, _
pseudo, lignes_contenu(), score_de_depart, monter, descendre








