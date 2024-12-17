' Créer les objets nécessaires
Set player = CreateObject("WMPlayer.OCX")
Set fso = CreateObject("Scripting.FileSystemObject")
Set voix = CreateObject("SAPI.SpVoice")
voix.Rate = 0 ' Vitesse normale
voix.Volume = 100 ' Volume au maximum


' Récupérer le chemin du dossier du script
scriptPath = fso.GetParentFolderName(WScript.ScriptFullName)
audioPath = scriptPath & "\musique.mp3" ' Nom du fichier audio

texte = "HE HE HE AAAAAAAAh!"
voix.Speak texte

' Configurer et jouer la musique en boucle
Set media = player.newMedia(audioPath)
player.currentPlaylist.appendItem(media)
player.settings.setMode "loop", True ' Activer la lecture en boucle
player.controls.play

' Fenêtre en boucle (impossible à fermer proprement)
Do
    MsgBox "Ba alors, tu t'est fait avoir", 16, "Alerte Système"
    
Loop
