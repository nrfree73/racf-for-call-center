#pragma compile(FileVersion, 2.0)
#pragma compile(ProductVersion, 2.0)
#pragma compile(FileDescription, [T41-DSEM v2.0])
#pragma compile(ProductName, [T41-DSEM v2.0])
#pragma compile(LegalCopyright, (01/2022) - Nicolas RISTOVSKI)
#pragma compile(CompanyName, Nicolas RISTOVSKI / EAPI69)

#include <file.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>

if  FileExists(@scriptdir & "\t41-DSEM.ini") Then
   FileMove(@scriptdir & "\t41-DSEM.ini",@scriptdir & "\t41 SVP-DSEM.ini")
   EndIf

if FileExists(@scriptdir & "\t41 SVP-DSEM.ini") Then
;$file=FileOpen(@scriptdir & "\t41.ini",0) ; 0= Read mode (default)
$file=@scriptdir & "\t41 SVP-DSEM.ini"
Else
$file=@scriptdir & "\t41 SVP-DSEM.ini"
   MsgBox(0,"Premier lancement ?","creation d'un nouveau fichier 't41 SVP-DSEM.ini'")
 IniWrite($file, "T41", "compte", "Identifiant;password")
 FileWriteLine($file,@crlf)
 IniWrite($file, "Website", "t41", "https://stmcv1.sf.intra.laposte.fr:444/ws_t41/T41/canalweb/login.ea" )
 FileWriteLine($file,@crlf)
 IniWrite($file, "Temporisation", "logon", "1500" )
 FileWriteLine($file,@crlf)
 IniWrite($file, "Rumba", "Macros", "C:\Users\" & StringStripWS(@username, 8) & "\AppData\Local\VirtualStore\Program Files (x86)\RUMBARH\System\Macro" )

;macros samples de base
DirCreate("C:\Users\" & StringStripWS(@username, 8) & "\AppData\Local\VirtualStore\Program Files (x86)\RUMBARH\Mframe\Macro")

;fenetres Rumba SVP DSEM
FileInstall( "C:\Appliloc\outils exe\macros t41\CICPFT00.WDM" , "C:\Users\" & StringStripWS(@username, 8) & "\AppData\Local\VirtualStore\Program Files (x86)\RUMBARH\System\Macro\" , 0) ; 0 = ne pas réecrire si existant..
FileInstall( "C:\Appliloc\outils exe\macros t41\SX11.WDM"     , "C:\Users\" & StringStripWS(@username, 8) & "\AppData\Local\VirtualStore\Program Files (x86)\RUMBARH\System\Macro\" , 0) ; 0 = ne pas réecrire si existant..
FileInstall( "C:\Appliloc\outils exe\macros t41\SX31.WDM"     , "C:\Users\" & StringStripWS(@username, 8) & "\AppData\Local\VirtualStore\Program Files (x86)\RUMBARH\System\Macro\" , 0) ; 0 = ne pas réecrire si existant..

;macros SVP DSEM
FileInstall( "C:\Appliloc\outils exe\macros t41\sx11.rmc"     , "C:\Users\" & StringStripWS(@username, 8) & "\AppData\Local\VirtualStore\Program Files (x86)\RUMBARH\System\Macro\" , 0)
FileInstall( "C:\Appliloc\outils exe\macros t41\sx31.rmc"     , "C:\Users\" & StringStripWS(@username, 8) & "\AppData\Local\VirtualStore\Program Files (x86)\RUMBARH\System\Macro\" , 0)
FileInstall( "C:\Appliloc\outils exe\macros t41\cicpft00.rmc" , "C:\Users\" & StringStripWS(@username, 8) & "\AppData\Local\VirtualStore\Program Files (x86)\RUMBARH\System\Macro\" , 0)
;
  MsgBox(0,"Information !","Changez les valeurs de votre compte dans le fichier généré 't41 SVP-DSEM.ini'" & @crlf & "avec le bloc-notes ... sauvegardez et puis relancez t41 SVP-DSEM.exe" & @crlf & "DSEM/EAPI69: nicolas RISTOVSKI => enjoy t41 SVP-DSEM.exe !" & @crlf & "fichier ini crée sous:" & @crlf & $file)
  Run("notepad.exe " & $file )
   Exit
   EndIf

;compte
$logon=IniRead($file,"T41","compte",0)
if  $logon<>"" Then ;lecture ok
$logon=stringsplit($logon,";") ;
$identifiant=$logon[1]
$mdp=$logon[2]

$lancert41= IniRead($file,"T41","lancer",1)
;MsgBox(0," ",$lancert41)


ElseIf $logon=0 Then ;IniRead ko car 0
   MsgBox(0,"Warning !","credentials non renseignés ? fichier t41 SVP-DSEM.ini")
   Exit
EndIf

$Website= IniRead($file,"Website","t41",0)
if $Website ='0' Then
   $Website = "https://stmcv1.sf.intra.laposte.fr:444/ws_t41/T41/canalweb/login.ea"
EndIf

$Temporisation= IniRead($file,"Temporisation","logon",1500)

$Rumba= IniRead($file,"Rumba","Macros",0)
if $Rumba ='0' Then
   $Rumba = "C:\Users\" & StringStripWS(@username, 8) & "\AppData\Local\VirtualStore\Program Files (x86)\RUMBARH\System\Macro"
EndIf

fileclose($file)


	 $fenetreRumba=""
	 ;choix fenetre Rumba avant execution...
	 ;==================================
	  Local $hGUI = GUICreate("Choix fenetre Rumba (SVP/DSEM)", 350, 50)
	  ; Crée un contrôle ComboBox.
    Local $idComboBox = GUICtrlCreateCombo("", 10, 10, 185, 20)
    ; Local $idFermer = GUICtrlCreateButton("Fermer", 210, 170, 85, 25)

    ; Ajoute des éléments supplémentaires à la ComboBox.
    GUICtrlSetData($idComboBox, "SX11|SX31|CICPFT00", "CICPFT00")

    ; Affiche la GUI.
    GUISetState(@SW_SHOW, $hGUI)

    Local $sComboRead = ""

    ; Boucle jusqu'à ce que l'utilisateur quitte.
    Local $idMsg = GUIGetMsg()
    While ($idMsg <> $GUI_EVENT_CLOSE) ; And ($idMsg <> $idFermer)
        If $idMsg = $idComboBox Then
            $fenetreRumba = GUICtrlRead($idComboBox)
         ;   MsgBox($MB_SYSTEMMODAL, "", "choisi: " & $fenetreRumba, 0, $hGUI)
        EndIf
        $idMsg = GUIGetMsg()
	 WEnd
	  $fenetreRumba = GUICtrlRead($idComboBox)
	 GUIDelete($hGUI)

	 if $fenetreRumba="" Then
		MsgBox(0, "Info !", "aucune fenetre Rumba choisie ! abandon..." & $fenetreRumba)
exit
		EndIf

;	 MsgBox(0, "Info !", "fenetre Rumba: " & $fenetreRumba & ".WDM")

	;===================================
  if $lancert41=1 then
$t= MsgBox(4,"Connexion " & $fenetreRumba &  " ?", "Voulez-vous lancer une connexion à t41 (edge) et Rumba (" & $fenetreRumba & ") ?" & @crlf & "	Assurez-vous que t41 (edge) soit deconnecté et fermé...		" & @crlf )
  Else
$t= MsgBox(4,"Connexion " & $fenetreRumba &  " ?", "Voulez-vous lancer une connexion à Rumba (" & $fenetreRumba & ") ?" & @crlf )
  EndIf


  If $t = 6 Then ;yes

;t41.. execution
if $lancert41=1 then ;lancer=1
 ShellExecute($Website)
;WinWaitActive("Title", "Intranet La Banque Postale - Authentification",15)
sleep($Temporisation) ; en ms...
tooltip("Envoi des credentials t41...",5,5,"")
;send credentials... t41
send($identifiant)
sleep(200)
send("{TAB}")
sleep(500)
send($mdp)
Sleep(200)
Send("{ENTER}")
sleep(500)
ToolTip("",5,5,"")
sleep($Temporisation)
EndIf


ToolTip ( "Lance la macro Rumba:  " & $fenetreRumba & ".rmc  ",5 ,5 ,"" )
	  ;MsgBox(0,"Executer la macro Rumba !","macro à executer:" & @crlf & @crlf & "   'Connexion DEX" & $Dex & ".rmc'   ")

 ;Rumba / RACF credentials from macros.. execution
 ShellExecute( $rumba & "\" & $fenetreRumba & ".WDM" ) ; "\Rumba-RACF.WDM"
sleep(500)
WinWaitActive("[CLASS:WdPageFrame]","",10)
sleep(2000)
send("{ALT}") ; fenetre dialogue ecran Rumba.Wdm pour choix macros
sleep(200)
Send("{O}")
sleep(200)
send("{x}")
;parcourir
sleep(200)

WinWaitActive("[CLASS:#32770]","",10)
sleep(200)
Send("{TAB}")
sleep(200)
;parcourir..
Send("{TAB}")
sleep(200)
send("{p}") ; correctif pour : Parcourir
;
Send($rumba & "\" & $fenetrerumba & ".rmc") ;path
;
Send("{TAB}")
sleep(200)
Send("{TAB}")
sleep(200)
;
Send("{v}");ouvrir
sleep(200)
;


;send($rumba & "\" & $fenetreRumba)
;sleep(200)

;send (  $fenetreRumba ) ; "Connexion " &
;sleep(200)
;
;Send("{TAB}")
;sleep(200)
;Send("{TAB}")
;sleep(200)
;
;Send("{ENTER}")
;sleep(200)

;
ToolTip("  t41 SVP-DSEM.exe   =>   DSEM/EAPI69/nicolas RISTOVSKI  (01/2022)",5,5,"")
sleep(5000)
ToolTip("",5,5,"")
EndIf ; $t=6 yes


