

# CustomProperties

L’objet `CustomProperties` représente les propriétés personnalisées qui sont propres à un élément particulier et à un complément de messagerie pour Outlook. Par exemple, il peut s’avérer nécessaire pour un complément de messagerie d’enregistrer certaines données propres au message électronique actif ayant activé le complément. Quand l’utilisateur consulte à nouveau le même message et réactive le complément de messagerie, ce dernier peut récupérer les données enregistrées en tant que propriétés personnalisées.

Étant donné qu’Outlook pour Mac ne met pas en cache les propriétés personnalisées, si le réseau de l’utilisateur tombe en panne, les compléments de messagerie ne peuvent pas accéder à leurs propriétés personnalisées.

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

### Exemple

L’exemple suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode [`saveAsync`](#saveasync) pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple utilise la méthode [`get`](CustomProperties.md#get) pour lire la propriété personnalisée `myProp`, utilise la méthode [`set`](CustomProperties.md#set) pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.

```JavaScript
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var mailbox = Office.context.mailbox;
    mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
  });
}
function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

### Méthodes

####  get(name) → {String}

Retourne la valeur de la propriété personnalisée spécifiée.

##### Paramètres :

|Nom| Type| Description|
|---|---|---|
|`name`| Chaîne|Nom de la propriété personnalisée à retourner.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### Renvoie :

Valeur de la propriété personnalisée spécifiée.

<dl class="param-type">

<dt>Type</dt>

<dd>Chaîne</dd>

</dl>

####  remove(name)

Supprime la propriété spécifiée de la collection de propriétés personnalisées.

Pour rendre la suppression de la propriété permanente, vous devez appeler la méthode [`saveAsync`](#saveasync) de l’objet `CustomProperties`.

##### Paramètres :

|Nom| Type| Description|
|---|---|---|
|`name`| Chaîne|Nom de la propriété à supprimer.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|
####  saveAsync([callback], [asyncContext])

Enregistre les propriétés personnalisées propres aux éléments sur le serveur.

Vous devez appeler la méthode `saveAsync` pour conserver les modifications effectuées avec la méthode [`set`](CustomProperties.md#set) ou la méthode [`remove`](#remove) de l’objet `CustomProperties`. L’enregistrement est une action asynchrone.

Il est recommandé de faire en sorte que la fonction de rappel vérifie et traite les erreurs provenant de `saveAsync`. Plus particulièrement, un complément de lecture peut être activé lorsque l’utilisateur est connecté dans un formulaire de lecture, puis l’utilisateur peut se déconnecter. Si le complément appelle `saveAsync` lorsqu’il est déconnecté, `saveAsync` renvoie une erreur. La méthode de rappel doit pouvoir gérer cette erreur.

##### Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`callback`| function| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult). |
|`asyncContext`| Object| &lt;facultatif&gt;|Toutes les données d’état transmises à la méthode de rappel.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### Exemple

L’exemple de code JavaScript suivant montre comment utiliser de manière asynchrone la méthode `loadCustomPropertiesAsync` pour charger des propriétés personnalisées propres à l’élément actif, ainsi que la méthode [`saveAsync`](saveasynccallback-asynccontext) pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode [`get`](CustomProperties.md#get) pour lire la propriété personnalisée `myProp`, utilise la méthode [`set`](CustomProperties.md#set) pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.

```JavaScript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
}
function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed){
    write(asyncResult.error.message);
  }
  else {
    // Async call to save custom properties completed.
    // Proceed to do the appropriate for your add-in.
  }
}

// Writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  set(name, value)

Affecte la valeur spécifiée à la propriété spécifiée.

La méthode `set` affecte la valeur spécifiée à la propriété spécifiée. Vous devez utiliser la méthode [`saveAsync`](#saveasync) pour enregistrer la propriété sur le serveur.

La méthode `set` crée une propriété si la propriété spécifiée n’existe pas déjà ; sinon, la valeur existante est remplacée par la nouvelle valeur. Le paramètre `value` peut être de n’importe quel type ; toutefois, il est toujours transmis au serveur sous forme de chaîne.

##### Paramètres :

|Nom| Type| Description|
|---|---|---|
|`name`| Chaîne|Nom de la propriété à définir.|
|`value`| Object|Valeur de la propriété à définir.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|