
# Obtenir et définir des métadonnées pour un complément Outlook

Vous pouvez gérer les données personnalisées dans votre complément Outlook en utilisant une des solutions suivantes :

- Les paramètres d’itinérance, qui permettent de gérer des données personnalisées pour la boîte aux lettres d’un utilisateur.
    
- Les propriétés personnalisées, qui permettent de gérer des données personnalisées pour un élément de boîte aux lettres d’un utilisateur.
    
Ces deux méthodes donnent accès aux données personnalisées auxquelles seul votre complément Outlook a accès, mais chaque méthode stocke les données de façon distincte. Autrement dit, les propriétés personnalisées n’ont pas accès aux données stockées par le biais des paramètres d’itinérance et inversement. Les données sont stockées sur le serveur de la boîte aux lettres et sont accessibles dans les sessions Outlook ultérieures sur tous les formats pris en charge par le complément. 

## Données personnalisées par boîte aux lettres : paramètres d’itinérance


Vous pouvez indiquer des données propres à la boîte aux lettres Exchange d’un utilisateur, à l’aide de l’objet [RoamingSettings](../../reference/outlook/RoamingSettings.md), telles que les préférences et les données personnelles de l’utilisateur. Votre complément de messagerie peut accéder aux paramètres d’itinérance lorsqu’il est en itinérance sur un appareil pour lequel il a été conçu (ordinateur, tablette ou smartphone).

 Les modifications apportées à ces données sont stockées dans une copie en mémoire de ces paramètres pour la session Outlook en cours. Vous devez explicitement enregistrer tous les paramètres d’itinérance après les avoir mis à jour afin qu’ils soient disponibles lors de la prochaine ouverture de votre complément, sur le même appareil ou sur un autre appareil pris en charge.


### Format des paramètres d’itinérance


Les données dans un objet  **RoamingSettings** sont stockées sous la forme d’une chaîne JSON (JavaScript Object Notation) sérialisée. L’exemple suivant illustre la structure, en partant du principe que trois paramètres d’itinérance sont définis et nommés `add-in_setting_name_0`,  `add-in_setting_name_1` et `add-in_setting_name_2`.


```js
{
  "add-in_setting_name_0":"add-in_setting_value_0",
  "add-in_setting_name_1":"add-in_setting_value_1",
  "add-in_setting_name_2":"add-in_setting_value_2"
}
```


### Chargement des paramètres d’itinérance


Un complément de messagerie charge généralement les paramètres d’itinérance dans le gestionnaire d’événements [Office.initialize](../../reference/shared/office.initialize.md). L’exemple de code JavaScript suivant montre comment charger des paramètres d’itinérance existants et obtenir la valeur de deux paramètres : « customerName » et « customerBalance ».


```js
var _mailbox;
var _settings;
var _customerName;
var _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}

```


### Création ou affectation d’un paramètre d’itinérance


Pour faire suite à l’exemple précédent, la fonction JavaScript suivante,  `setAddInSetting`, illustre l’utilisation de la méthode [RoamingSettings.set](../../reference/outlook/RoamingSettings.md) pour régler un paramètre nommé `cookie` à la date d’aujourd’hui, et rendre les données persistantes en utilisant la méthode [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md) pour réenregistrer tous les paramètres d’itinérance sur le serveur. La méthode **set** crée le paramètre si celui-ci n’existe pas déjà, et affecte au paramètre la valeur spécifiée. La méthode **saveAsync** enregistre les paramètres d’itinérance en mode asynchrone. Cet exemple de code passe une méthode de rappel, `saveMyAddInSettingsCallback`, à  **saveAsync**. Lorsque l’appel asynchrone se termine,  `saveMyAddInSettingsCallback` est appelé à l’aide d’un paramètre, _asyncResult_. Ce paramètre est un objet [AsyncResult](../../reference/outlook/simple-types.md) qui contient les résultats de l’appel asynchrone et de tous les détails le concernant. Vous pouvez utiliser le paramètre facultatif _userContext_ pour passer des informations d’état de l’appel asynchrone à la fonction de rappel.


```js
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback method after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


### Suppression d’un paramètre d’itinérance


Toujours dans le prolongement des exemples précédents, la fonction JavaScript suivante,  `removeAddInSetting`, illustre l’utilisation de la méthode [RoamingSettings.remove](../../reference/outlook/RoamingSettings.md) pour supprimer le paramètre `cookie` et réenregistrer tous les paramètres d’itinérance sur le serveur Exchange.


```js
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```


## Données personnalisées par élément dans une boîte aux lettres : propriétés personnalisées


Vous pouvez spécifier les données propres à un élément dans la boîte aux lettres de l’utilisateur à l’aide de l’objet [CustomProperties](../../reference/outlook/CustomProperties.md). Par exemple, votre complément de messagerie peut catégoriser certains messages et noter la catégorie à l’aide d’une propriété personnalisée  `messageCategory`. Si votre complément de messagerie crée des rendez-vous à partir de suggestions de réunion dans un message, vous pouvez utiliser une propriété personnalisée pour suivre chacun de ces rendez-vous. Cela garantit que si l’utilisateur ouvre à nouveau le message, votre complément de messagerie ne propose pas de créer le rendez-vous une seconde fois.

Comme pour les paramètres d’itinérance, les modifications apportées aux propriétés personnalisées sont stockées dans des copies en mémoire des propriétés de la session Outlook en cours. Pour vous assurer que les propriétés personnalisées seront disponibles dans la prochaine session, enregistrez-les toutes sur le serveur.

Ces propriétés personnalisées propres à un complément et à un élément sont accessibles uniquement à l’aide de l’objet  **CustomProperties**. Ces propriétés sont différentes des propriétés MAPI personnalisées [UserProperties](http://msdn.microsoft.com/library/20b49c86-d74f-9bda-382c-559af278c148%28Office.15%29.aspx) dans le modèle objet Outlook, et des propriétés étendues dans Services Web Exchange (EWS). Vous ne pouvez pas accéder à **CustomProperties** en utilisant le modèle objet Outlook ou EWS.

Cependant, un complément de messagerie peut obtenir des propriétés étendues MAPI à l’aide de l’opération [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx). Accédez à  **GetItem** côté serveur en utilisant un jeton de rappel, ou côté client en utilisant la méthode [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md). Dans la demande  **GetItem**, indiquez les propriétés étendues personnalisées dont vous avez besoin dans un ensemble de propriétés. Un complément de messagerie peut également utiliser  **makeEwsRequestAsync**, ainsi que les opérations [CreateItem](http://msdn.microsoft.com/library/78a52120-f1d0-4ed7-8748-436e554f75b6%28Office.15%29.aspx) et [UpdateItem](http://msdn.microsoft.com/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) EWS pour créer et modifier les propriétés étendues.




### Utilisation de propriétés personnalisées


Avant de pouvoir utiliser des propriétés personnalisées, vous devez les charger en appelant la méthode [loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md). Si des propriétés personnalisées sont déjà définies pour l’élément actif, elles sont chargées depuis le serveur Exchange à ce stade. Une fois que vous avez créé le conteneur de propriétés, vous pouvez utiliser les méthodes [set](../../reference/outlook/CustomProperties.md) et [get](../../reference/outlook/CustomProperties.md) pour ajouter et récupérer des propriétés personnalisées. Pour enregistrer les modifications que vous apportez au conteneur de propriétés, vous devez utiliser la méthode [saveAsync](../../reference/outlook/CustomProperties.md) pour conserver les modifications sur le serveur Exchange.


 >**Remarque**  Comme Outlook pour Mac ne met pas en cache les propriétés personnalisées, si le réseau de l’utilisateur tombe en panne, les compléments de messagerie dans Outlook pour Mac ne seront pas en mesure d’accéder à leurs propriétés personnalisées.


### Exemple de propriétés personnalisées


L’exemple suivant illustre un ensemble simplifié des méthodes pour un complément Outlook qui utilise des propriétés personnalisées. Vous pouvez utiliser cet exemple comme point de départ pour votre complément qui utilise des propriétés personnalisées. 

Cet exemple inclut les méthodes suivantes :


- [Office.initialize](../../reference/shared/office.initialize.md) -- Initialise le complément et charge le conteneur de propriétés personnalisées depuis le serveur Exchange.
    
-  **customPropsCallback** -- Obtient le conteneur de propriétés personnalisées qui est renvoyé depuis le serveur et l’enregistre pour une utilisation ultérieure.
    
-  **updateProperty** -- Définit ou met à jour une propriété spécifique, puis enregistre la modification sur le serveur.
    
-  **removeProperty** -- Supprime une propriété spécifique à partir du conteneur de propriétés, puis enregistre la suppression sur le serveur.
    



```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
}

// Callback function from loading custom properties.
function customPropsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
  else {
    // Successfully loaded custom properties,
    // can get them from the asyncResult argument.
    _customProps = asyncResult.value;
  }
}

// Get individual custom property.
function getProperty() {
  var myProp = customProps.get("myProp");
}

// Set individual custom property.
function updateProperty(name, value) {
  _customProps.set(name, value);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Remove a custom property.
function removeProperty(name) {
  _customProps.remove(name);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Callback function from saving custom properties.
function saveCallback() {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


## Ressources supplémentaires

    
- [Vue d'ensemble de la propriété MAPI](http://msdn.microsoft.com/library/02e5b23f-1bdb-4fbf-a27d-e3301a359573%28Office.15%29.aspx)
    
- [Présentation des propriétés Outlook](http://msdn.microsoft.com/library/242c9e89-a0c5-ff89-0d2a-410bd42a3461%28Office.15%29.aspx)
    
- [Appeler des services web à partir d’un complément Outlook](../outlook/web-services.md)
    
- [Les propriétés et les propriétés étendues dans EWS dans Exchange](http://msdn.microsoft.com/library/68623048-060e-4602-b3fa-62617a94cf72%28Office.15%29.aspx)
    
- [Jeux de propriétés et de réponse des formes dans EWS dans Exchange](http://msdn.microsoft.com/library/04a29804-6067-48e7-9f5c-534e253a230e%28Office.15%29.aspx)
    


