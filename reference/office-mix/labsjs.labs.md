
# LabsJS.Labs

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Le module LabsJS.Labs contient l’ensemble d’interfaces API JavaScript clés que vous pouvez utiliser pour créer des compléments Office (les ateliers). Les API fournissent le point d’entrée pour développer des ateliers.

## Module d’API LabsJS.Labs

Le module Ateliers contient les types suivants :


### Variables


|||
|:-----|:-----|
|[Labs.DefaultHostBuilder](../../reference/office-mix/labs.defaulthostbuilder.md)|Utilisez cet objet pour créer une instance [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md) par défaut.|

### Fonctions


|||
|:-----|:-----|
|[Labs.Connect](../../reference/office-mix/labs.connect.md)|Initialise une connexion avec l’hôte.|
|[Labs.Connect (surcharge)](../../reference/office-mix/labs.connect-overload.md)|Initialise une connexion avec l’hôte et fournit des paramètres d’entrée.|
|[Labs.isConnected](../../reference/office-mix/labs.isconnected.md)|Initialise une connexion avec l’hôte.|
|[Labs.getConnectionInfo](../../reference/office-mix/labs.getconnectioninfo.md)|Récupère des informations de configuration associées à une connexion spécifiée.|
|[Labs.disconnect](../../reference/office-mix/labs.disconnect.md)|Déconnecte l’atelier de l’hôte et indique que l’atelier est terminé.|
|[Labs.editLab](../../reference/office-mix/labs.editlab.md)|Ouvre l’atelier spécifié pour le modifier. Il est possible d’indiquer des données de configuration de l’atelier en mode Édition. Toutefois, il est impossible de modifier un atelier lors de son exécution.|
|[Labs.takeLab](../../reference/office-mix/labs.takelab.md)|Exécute l’atelier spécifié et active l’envoi de résultats de l’atelier au serveur. Un atelier ne peut pas être exécuté lorsqu’il est en cours de modification.|
|[Labs.on](../../reference/office-mix/labs.on.md)|Ajoute un nouveau gestionnaire pour un événement spécifié.|
|[Labs.off](../../reference/office-mix/labs.off.md)|Supprime un gestionnaire d’événements pour un événement spécifié.|
|[Labs.getTimeline](../../reference/office-mix/labs.gettimeline.md)|Récupère une instance d’objet [Labs.Timeline](../../reference/office-mix/labs.timeline.md) qui peut être utilisée pour commander le contrôle de lecteur d’hôte.|
|[Labs.registerDeserializer](../../reference/office-mix/labs.registerdeserializer.md)|Désérialise un objet JSON spécifié en un objet. Seuls les auteurs de composant doivent l’utiliser.|

### Classes


|||
|:-----|:-----|
|[Labs.ComponentInstanceBase](../../reference/office-mix/labs.componentinstancebase.md)|Classe de base pour l’initialisation d’instances de composant.|
|[Labs.ComponentInstance](../../reference/office-mix/labs.componentinstance.md)|Représente l’instance d’un composant, qui est une instanciation d’un composant donné pour un utilisateur lors de l’exécution. L’objet comporte une vue traduite du composant pour une exécution spécifique de l’atelier.|
|[Labs.Command](../../reference/office-mix/labs.command.md)|Commande générale permettant de transmettre des messages entre le client et l’hôte.|
|[Labs.LabEditor](../../reference/office-mix/labs.labeditor.md)|L’objet **LabEditor** vous permet de modifier un atelier donné, ainsi que d’obtenir et de définir des données de configuration associées à l’atelier.|
|[Labs.LabInstance](../../reference/office-mix/labs.labinstance.md)|Instance d’un atelier configurée pour l’utilisateur actuel. Cet objet permet d’enregistrer et de récupérer des données relatives à l’atelier pour l’utilisateur.|
|[Labs.Timeline](../../reference/office-mix/labs.timeline.md)|Fournit un accès à la fonctionnalité de chronologie labs.js.|
|[Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md)|Objet conteneur comportant et suivant des valeurs pour un atelier spécifié. Les valeurs peuvent être stockées localement ou sur le serveur.|

### Interfaces


|||
|:-----|:-----|
|[Labs.GetActionsCommandData](../../reference/office-mix/labs.getactionscommanddata.md)|Permet de récupérer des données associées à la commande [LabsJS.Labs.Core.GetActions](../../reference/office-mix/labsjs.labs.core.getactions.md).|
|[Labs.IMessageHandler](../../reference/office-mix/labs.imessagehandler.md)|Interface permettant de définir des gestionnaires d’événements.|
|[Labs.ITimelineNextMessage](../../reference/office-mix/labs.itimelinenextmessage.md)|Permet d’interagir avec l’objet [Labs.Core.IMessage](https://msdn.microsoft.com/library/office/mt599680.aspx).|
|[Labs.SendMessageCommandData](../../reference/office-mix/labs.sendmessagecommanddata.md)|Données associées à une commande [Labs.CommandType.TakeAction](https://msdn.microsoft.com/library/office/mt599680.aspx).|
|[Labs.TakeActionCommandData](../../reference/office-mix/labs.takeactioncommanddata.md)|Données associées à une commande de prise d’action.|

### Énumérations


|||
|:-----|:-----|
|[Labs.ConnectionState](../../reference/office-mix/labs.connectionstate.md)|Énumère les états de connexion possibles entre l’atelier et l’hôte.|
|[Labs.ProblemState](../../reference/office-mix/labs.problemstate.md)|Valeurs d’état pour un atelier donné.|
