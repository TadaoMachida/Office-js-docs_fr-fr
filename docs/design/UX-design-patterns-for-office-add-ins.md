# Modèles de conception de l’expérience utilisateur pour les compléments Office. 

Quand vous créez des compléments Office, vous devez concevoir une expérience utilisateur intéressante qui étend les possibilités d’Office. Pour créer un complément idéal, votre complément doit, entre autres, offrir une première expérience intéressante aux utilisateurs et assurer des transitions fluides entre les pages. En offrant aux utilisateurs une expérience nette et moderne, vous les persuaderez de continuer à utiliser votre complément. Cet article présente les ressources liées à l’expérience utilisateur pouvant être utilisées par les développeurs et les concepteurs qui :

* Décrivent des modèles de conception d’expérience utilisateur courants en faisant appel aux meilleures pratiques.
* Implémentent des styles et des composants de la structure Office.
* Implémentent des compléments ressemblant à une extension naturelle de l’interface utilisateur d’Office par défaut. 

## Commencer à utiliser les exemples de ressources pour concevoir des compléments Office

L’utilisation de ces éléments de conception ou de code ne demande aucun prérequis. Pour créer une expérience utilisateur parfaite pour votre complément, procédez comme suit :

* Passez en revue les modèles de conception d’expérience utilisateur et identifiez les modèles importants pour votre complément. Par exemple, sélectionnez une des premières expériences d’utilisation.
* Puis, effectuez une ou plusieurs des actions suivantes :
	* Copiez les fichiers de code dans votre projet de complément et commencez à le personnaliser pour répondre à vos besoins. Vous devez avoir le [fichier common.js](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/), le [dossier Assets](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/assets) et le dossier Code pour choisir le modèle de conception dont vous avez besoin. Consultez les liens ci-dessous.
	* Téléchargez les fichiers PDF de référence et utilisez-les pour concevoir votre propre expérience utilisateur. Consultez les liens ci-dessous.
	* Téléchargez les fichiers Adobe Illustrator et modifiez-les pour imiter vos propres modèles de complément. Obtenez-les [ici](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Source%20Files).
 

## Première expérience d’utilisation

Il s’agit de l’expérience vécue par un utilisateur lorsqu’il ouvre votre complément pour la première fois. Les points suivants répertorient les modèles de conception à intégrer pour la première exécution de votre complément. Vous trouverez une image de chacun d’entre eux en dessous.

* **Étapes pour commencer** permet aux utilisateurs ayant une liste d’étapes à suivre de commencer à utiliser votre complément. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_StepsToStart.pdf "PDF"), [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/instruction-step))
* **Valeur** communique la proposition de valeur de votre complément. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_ValuePlacemat.pdf "PDF"), [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/value-placemat))
* **Vidéo** présente aux utilisateurs une vidéo avant qu’ils commencent à utiliser votre complément. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_VideoPlacemat.pdf "PDF"), [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/video-placemat))
* **Procédure pas à pas** présente aux utilisateurs une série de fonctionnalités ou d’informations avant qu’ils commencent à utiliser le complément. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_PagingPanel.pdf "PDF"), [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/walkthrough))
* L’[Office Store](https://msdn.microsoft.com/fr-fr/library/office/jj220033.aspx) dispose d’un système destiné à fournir aux utilisateurs la version d’évaluation d’un complément. Cependant, si vous souhaitez contrôler l’interface utilisateur pendant un essai, utilisez les modèles suivants :
	* **Version d’évaluation** présente aux utilisateurs comment utiliser la version d’évaluation de votre complément. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/FirstRun_TrialVersion.pdf "PDF"), [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/trial-placemat))
	* **Fonctionnalité d’évaluation** informe les utilisateurs que la fonctionnalité qu’ils tentent d’utiliser n’est pas disponible dans la version d’évaluation du complément. ([code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/first-run/trial-placemat-feature))


> Remarque : Déterminez s’il convient de montrer la vidéo sur la première expérience d’utilisation une ou plusieurs fois (tout dépend de son importance pour votre scénario). Par exemple, si les utilisateurs utilisent régulièrement votre complément, ils peuvent ne plus se rappeler de la façon de l’utiliser. Pour ces utilisateurs, il peut être utile de revoir cette première expérience d’utilisation. 

 <table>
 <tr><th>Étapes pour commencer</th><th>Valeur</th><th>Vidéo</th></tr>
 <tr><td>![instruction steps" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/instruction.step.PNG)</td><td>![value placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/value.placemat.PNG)</td><td>![video placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/video.placemat.PNG)</td></tr>
 </table>

 <table>
 <tr><th>Première page de la procédure pas à pas</th><th>Version d’évaluation</th><th>Fonctionnalité de la version d’évaluation</th></tr>
 <tr><td>![walkthrough 1" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/walkthrough1.PNG)</td><td>![trial placemat" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/trial.placemat.PNG)</td><td>![trial placemat feature" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/trial.placemat.feature.PNG)</td></tr>
 </table> 


## Générique et personnalisation

* **Page d’accueil** est l’endroit où se rendent les utilisateurs une fois la première expérience d’utilisation ou le processus de connexion terminés. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Helpful%20Templates/AddIn_Template_Standard_Layout.pdf "PDF"), [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/generic/landing-page))

<table>
 <tr><th>Accueil</th></tr>
 <tr><td>![landing page" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/landing.page.PNG)</td></tr>
 </table>

## Notifications

Votre complément peut avertir les utilisateurs d’un événement (une erreur, par exemple) ou de l’état d’avancement d’un élément de plusieurs façons. Les points suivants répertorient ces méthodes. Vous trouverez une image de chacun d’entre eux en dessous.

* **Boîte de dialogue incorporée**  affiche une boîte de dialogue dans le volet Office qui vous fournit des informations et, éventuellement, une expérience interactive, à l’aide des boutons ou d’autres commandes. Pensez à en utiliser une pour inviter un utilisateur à confirmer une action. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Embedded_Dialog.pdf "PDF") , [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/embedded-dialog))
* **Message incorporé** indique l’échec, la réussite ou des informations et peut apparaître à un emplacement spécifié dans le volet Office. Par exemple, si un utilisateur entre une adresse de messagerie erronée dans une zone de texte, un message d’erreur apparaît juste en dessous de la zone de texte. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_Inline_Message.pdf "PDF"), [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/inline-message))
* **Bannière de message** fournit des informations et, éventuellement, des instructions dans une bannière qui peut être réduite à une seule ligne, développée en plusieurs lignes ou masquée. Pensez à utiliser des bannières de message pour signaler une mise à jour du service ou donner un conseil utile lorsque le complément démarre. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_messagebanner.pdf "PDF"), [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/message-banner))
* **Barre de progression** indique la progression d’un processus long et synchrone, tel qu’une tâche de configuration qui doit être terminée pour que l’utilisateur puisse effectuer d’autres actions. Il s’agit d’une page distincte interstitielle qui met en évidence la marque du complément. Utilisez une barre de progression quand le processus peut envoyer des notifications pour indiquer la progression de la tâche dans le complément. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_progress.pdf "PDF"), [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/progress-bar))
* **Bouton fléché** indique qu’un processus synchrone long est lancé, mais ne fournit aucune indication sur son état d’avancement. Il s’agit d’une page distincte interstitielle qui met en évidence la marque du complément. Utilisez un bouton fléché quand le complément ne peut pas indiquer avec précision la progression du processus. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_progress.pdf "PDF"), [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/spinner))
* **Annonce** fournit un bref message qui disparaît au bout de quelques secondes. Comme il se peut que l’utilisateur ne voie pas le message, utilisez une annonce uniquement pour les informations non importantes. Utilisez une annonce pour informer les utilisateurs d’un événement dans un système distant, tel que la réception d’un message électronique. ([PDF](https://github.com/OfficeDev/Office-Add-in-Design-Patterns/blob/master/Patterns/Notification_toast.pdf "PDF"), [code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/tree/master/templates/notifications/toast))

 <table>
 <tr><th>Boîte de dialogue incorporée</th><th>Message incorporé</th><th>Bannière de message</th></tr>
 <tr><td>![embedded dialog" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/embedded.dialog.PNG)</td><td>![inline message" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/inline.message.PNG)</td><td>![message banner" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/message.banner.PNG)</td></tr>
 </table>

 <table>
 <tr><th>Barre de progression</th><th>Bouton fléché</th><th>Annonce</th></tr>
 <tr><td>![progress bar" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/progress.bar.PNG)</td><td>![spinner" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/spinner.PNG)</td><td>![toast" style="width: 264px;](https://github.com/OfficeDev/Office-Add-in-UX-Design-Pattern-Code/blob/master/Images/toast.PNG)</td></tr>
 </table>

## Problèmes connus

* L’exécution de certains fichiers de code en dehors d’un projet de complément génère une erreur JavaScript. 
	* Résolution : veillez à ajouter ces fichiers à un projet de complément Office. 
	
## Ressources supplémentaires

* [Meilleures pratiques de développement de compléments Office](https://dev.office.com/docs/add-ins/design/add-in-development-best-practices)
* [Structure de l’interface utilisateur Office](http://dev.office.com/fabric/)
