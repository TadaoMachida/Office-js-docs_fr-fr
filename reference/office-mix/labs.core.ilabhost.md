
# Labs.Core.ILabHost

 _**S’applique à :** applications pour Office |Compléments Office | Office Mix | PowerPoint_

Fournit une couche d’abstraction pour connecter Labs.js à l’hôte.

```
interface ILabHost
```


## Méthodes


### getSupportedVersions

 `getSupportedVersions(): Core.ILabHostVersionInfo[]`

Récupère les versions prises en charge par l’hôte de l’atelier.

 **Paramètres**

Aucun.


### connect

 `connect(versions: Core.ILabHostVersionInfo[], callback: Core.ILabCallback<Core.IConnectionResponse>)`

Initialise une connexion avec l’hôte.

 **Paramètres**


|||
|:-----|:-----|
| _versions_|Liste des versions hôtes que le client peut utiliser.|
| _callback_|Fonction de rappel qui se déclenche une fois la connexion établie.|

### disconnect

 `disconnect(callback: Core.ILabCallback<void>)`

Met fin à la communication avec l’hôte.

 **Paramètres**


|||
|:-----|:-----|
| _completionStatus_|État de l’atelier lors de la déconnexion.|
| _callback_|Fonction de rappel qui se déclenche une fois la déconnexion terminée.|

### actif

 `on(handler: (string: any, any: any): void)`

Ajoute un gestionnaire d’événements pour gérer les messages provenant de l’hôte. La promesse résolue est renvoyée à l’hôte.

 **Paramètres**


|||
|:-----|:-----|
| _handler_|Gestionnaire d’événements.|

### sendMessage

 `sendMessage(type: string, options: Core.IMessage, callback: Core.ILabCallback<Core.IMessageResponse>)`

Envoie un message à l’hôte.

 **Paramètres**


|||
|:-----|:-----|
| _type_|Type de message envoyé.|
| _options_|Options des messages.|
| _callback_|Fonction de rappel qui se déclenche une fois le message reçu.|

### rapidement

 `create(options: Core.ILabCreationOptions, callback: Core.ILabCallback<void>)`

Crée l’atelier. Stocke les informations de l’hôte et prévoit de l’espace pour stocker la configuration et d’autres éléments.

 **Paramètres**


|||
|:-----|:-----|
| _options_|Options transmises dans le cadre de l’opération Création.|
| _callback_|Fonction de rappel qui se déclenche une fois l’atelier créé.|

### getConfiguration

 `getConfiguration(callback: Core.ILabCallback<Core.IConfiguration>)`

Récupère la configuration de l’atelier en cours depuis l’hôte.

 **Paramètres**


|||
|:-----|:-----|
| _callback_|Fonction de rappel pour récupérer les informations de configuration.|

### setConfiguration

 `setConfiguration(configuration: Core.IConfiguration, callback: Core.ILabCallback<void>)`

Définit une nouvelle configuration pour l’atelier sur l’hôte.

 **Paramètres**


|||
|:-----|:-----|
| _configuration_|Configuration de l’atelier définie.|
| _callback_|Fonction de rappel qui se déclenche une fois la configuration définie.|

### getConfigurationInstance

 `getConfigurationInstance(callback: Core.ILabCallback<Core.IConfigurationInstance>)`

Récupère la configuration de l’instance pour l’atelier.

 **Paramètres**


|||
|:-----|:-----|
| _callback_|Fonction de rappel qui se déclenche une fois l’instance de configuration récupérée.|

### getState

 `getState(callback: Core.ILabCallback<any>)`

Récupère l’état actuel de l’atelier pour un utilisateur donné.

 **Paramètres**


|||
|:-----|:-----|
| _completionStatus_|Fonction de rappel qui renvoie l’état actuel de l’atelier.|

### setState

 `setState(state: any, callback: Core.ILabCallback<void>)`

Définit l’état de l’atelier pour un utilisateur donné.

 **Paramètres**


|||
|:-----|:-----|
| _state_|État de l’atelier.|
| _callback_|Fonction de rappel qui se déclenche une fois l’état défini.|

### takeAction

 `takeAction(type: string, options: Core.IActionOptions, callback: Core.ILabCallback<Core.IAction>)`

Tente une action.

 **Paramètres**


|||
|:-----|:-----|
| _type_|Type d’action.|
| _options_|Options fournies avec l’action.|
| _callback_|Fonction de rappel qui renvoie la dernière action exécutée.|

### takeAction

 `takeAction(type: string, options: Core.IActionOptions, result: Core.IActionResult, callback: Core.ILabCallback<Core.IAction>)`

Exécute une action qui a déjà été effectuée.

 **Paramètres**


|||
|:-----|:-----|
| _type_|Type d’action.|
| _options_|Options fournies avec l’action.|
| _result_|Résultat de l’action.|
| _callback_|Fonction de rappel qui renvoie la dernière action exécutée.|

### getActions

 `getActions(type: string, options: Core.IGetActionOptions, callback: Core.ILabCallback<Core.IAction[]>)`

Tente une action.

 **Paramètres**


|||
|:-----|:-----|
| _type_|Type d’action Get.|
| _options_|Options fournies avec l’action Get.|
| _callback_|Fonction de rappel qui renvoie la liste des actions effectuées.|
