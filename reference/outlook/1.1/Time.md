

# Heure

L’objet `Time` est renvoyé comme propriété [`start`](Office.context.mailbox.item.md#start-datetime) ou [`end`](Office.context.mailbox.item.md#end-datetime) d’un rendez-vous en mode composition.

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition|

### Méthodes

####  getAsync([options], callback)

Obtient l’heure de début ou de fin d’un rendez-vous.

##### Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult).

La date et l’heure sont fournies sous la forme d’un objet Date dans la propriété `asyncResult.value`. La valeur est exprimée au format UTC (temps universel coordonné). Vous pouvez convertir l’heure UTC au format du client avec la méthode [`convertToLocalClientTime`](Office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime).|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition|
####  setAsync(dateTime, [options], [callback])

Définit l’heure de début ou de fin d’un rendez-vous.

Si la méthode `setAsync` est appelée dans la propriété [`start`](Office.context.mailbox.item.md#start-datetime), la propriété [`end`](Office.context.mailbox.item.md#end-datetime) est modifiée pour maintenir la durée du rendez-vous telle que définie précédemment. Si la méthode `setAsync` est appelée dans la propriété `end`, la durée du rendez-vous est étendue jusqu’à la nouvelle heure de fin.

La durée doit être exprimée au format UTC. Vous pouvez obtenir l’heure UTC correcte à l’aide de la méthode [`convertToUtcClientTime`](Office.context.mailbox.md#converttoutcclienttimeinput--date).

##### Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`dateTime`| Date||Objet Date exprimé au format UTC (temps universel coordonné).|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult). <br/>Si la définition de la date et de l’heure échoue, la propriété `asyncResult.error` contient un code d’erreur.<br/><table class="nested-table"><thead><tr><th>Code d'erreur</th><th>Description</th></tr></thead><tbody><tr><td><code>InvalidEndTime</code></td><td>L’heure de fin du rendez-vous est antérieure à l’heure de début du rendez-vous.</td></tr></tbody></table>|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Mode Outlook applicable| Composition|

##### Exemple

L’exemple suivant définit l’heure de début d’un rendez-vous.

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
     asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```
