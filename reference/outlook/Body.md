

# Corps

L’objet `body` fournit des méthodes d’ajout et de mise à jour du contenu du message ou du rendez-vous. Il est renvoyé dans la propriété `body` de l’élément sélectionné.

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

### Méthodes

####  getAsync(coercionType, [options], [callback])

Renvoie le corps actif dans un format spécifié.

Cette méthode renvoie la totalité du corps actif dans le format spécifié par `coercionType`.

##### Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`coercionType`| [Office.CoercionType](Office.md#coerciontype-string)||Format du corps renvoyé.|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult).

Le corps est fourni au format demandé dans la propriété `asyncResult.value`.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.3|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### Exemples

Cet exemple obtient le corps du message en texte brut.

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext:"This is passed to the callback" },
  function callback(result) {
    // Do something with the result
  });
```

Voici un exemple du paramètre `result` transmis à la fonction de rappel.

```js
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  getTypeAsync([options], [callback])

Obtient une valeur qui indique si le contenu est au format HTML ou texte.

##### Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult).

Le type de contenu est renvoyé sous la forme d’une des valeurs [CoercionType](Office.md#coerciontype-string) dans la propriété `asyncResult.value`.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition|
####  prependAsync(data, [options], [callback])

Ajoute le contenu spécifié au début du corps de l’élément.

La méthode `prependAsync` insère la chaîne spécifiée au début du corps de l’élément. Appeler la méthode `prependAsync` revient à appeler la méthode [`setSelectedDataAsync`](#setselecteddataasync) avec le point d’insertion au début du contenu du corps.

Lorsque vous incluez des liens dans la balise HTML, vous pouvez désactiver l’aperçu du lien en ligne en définissant l’attribut `id` de l’ancre (`<a>`) sur `LPNoLP`. Par exemple :

```js
Office.context.mailbox.item.body.prependAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`data`| String||La chaîne doit être insérée au début du corps. Elle est limitée à un million de caractères.|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>Format du corps souhaité. La chaîne du paramètre <code>data</code> est convertie dans ce format.</td></tr><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult). <br/>Les erreurs rencontrées seront indiquées dans la propriété `asyncResult.error`.<br/><table class="nested-table"><thead><tr><th>Code d'erreur</th><th>Description</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>Le paramètre <code>data</code> comprend plus de 1 000 000 de caractères.</td></tr></tbody></table>|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Mode Outlook applicable| Composition|
####  setAsync (data, [options], [callback])

Remplace l’ensemble du corps avec le texte spécifié.

La méthode `setAsync` remplace le corps existant de l’élément par la chaîne spécifiée ou, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné.

Lorsque vous incluez des liens dans la balise HTML, vous pouvez désactiver l’aperçu du lien en ligne en définissant l’attribut `id` de l’ancre (`<a>`) sur `LPNoLP`. Par exemple :

```js
Office.context.mailbox.item.body.setAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`data`| String||Chaîne qui remplace le corps existant. Elle est limitée à 1 000 000 caractères.|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>Format du corps souhaité. La chaîne du paramètre <code>data</code> est convertie dans ce format.</td></tr><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult). <br/>Les erreurs rencontrées seront indiquées dans la propriété `asyncResult.error`.<br/><table class="nested-table"><thead><tr><th>Code d'erreur</th><th>Description</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>Le paramètre <code>data</code> comprend plus de 1 000 000 de caractères.</td></tr><tr><td><code>InvalidFormatError</code></td><td>Le paramètre <code>options.coercionType</code> est défini sur <code>Office.CoercionType.Html</code> et le corps du message est en texte brut.</td></tr></tbody></table>|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.3|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Mode Outlook applicable| Composition|

##### Exemples

L’exemple suivant remplace le corps par du contenu HTML.

```js
Office.context.mailbox.item.body.setAsync(
  "<b>(replaces all body, including threads you are replying to that may be on the bottom)</b>",
  { coercionType:"html", asyncContext:"This is passed to the callback" },
  function callback(result) {
    // Process the result
  });
```

Voici un exemple du paramètre `result` transmis à la fonction de rappel.

```js
{
  "value":null,
  "status":"succeeded",
  "asyncContext":"This is passed to the callback"
}
```

####  setSelectedDataAsync(data, [options], [callback])

Remplace la sélection dans le corps par le texte spécifié.

La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps de l’élément ou, si du texte est sélectionné dans l’éditeur, elle remplace ce texte. Si le curseur ne s’est jamais trouvé dans le corps de l’élément, ou si le corps de l’élément n’est plus la partie active de l’interface utilisateur, la chaîne est insérée au début du corps de l’élément.

Lorsque vous incluez des liens dans la balise HTML, vous pouvez désactiver l’aperçu du lien en ligne en définissant l’attribut `id` de l’ancre (`<a>`) sur `LPNoLP`. Par exemple :

```js
Office.context.mailbox.item.body.setSelectedDataAsync(
  '<a id="LPNoLP" href="http://www.contoso.com">Click here!</a>',
  {coercionType: Office.CoercionType.Html},
  callback);
```

##### Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`data`| String||Chaîne à insérer dans le corps. Elle est limitée à un million de caractères.|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>coercionType</code></td><td><a href="Office.md#coerciontype-string">Office.CoercionType</a></td><td>&lt;optional&gt;</td><td>Format du corps souhaité. La chaîne du paramètre <code>data</code> est convertie dans ce format.</td></tr><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult). <br/>Les erreurs rencontrées seront indiquées dans la propriété `asyncResult.error`.<br/><table class="nested-table"><thead><tr><th>Code d'erreur</th><th>Description</th></tr></thead><tbody><tr><td><code>DataExceedsMaximumSize</code></td><td>Le paramètre <code>data</code> comprend plus de 1 000 000 de caractères.</td></tr><tr><td><code>InvalidFormatError</code></td><td>Le type de corps est défini en mode HTML et le paramètre de données contient du texte brut.</td></tr></tbody></table>|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadWriteItem|
|Mode Outlook applicable| Composition|
