

# NotificationMessages

## NotificationMessages

L’objet `NotificationMessages` est renvoyé en tant que propriété [`notificationMessages`](Office.context.mailbox.item.md#notificationmessages-notificationmessages) d’un élément.

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.3|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

### Méthodes

####  addAsync(key, JSONmessage, [options], [callback])

Ajoute une notification à un élément.

Chaque message est limité à 5 notifications. Si vous en définissez plus, une erreur `NumberOfNotificationMessagesExceeded` est renvoyée.

##### Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`key`| String||Clé spécifiée par un développeur pour référencer ce message de notification. Les développeurs peuvent l’utiliser pour modifier ce message ultérieurement. Sa longueur ne peut pas être supérieure à 32 caractères.|
|`JSONmessage`| Objet||Objet JSON qui contient le message de notification à ajouter à l’élément. Il se compose des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Description</th></tr></thead><tbody><tr><td><code>type</code></td><td><a href="Office.MailboxEnums.md#.ItemNotificationMessageType">Office.MailboxEnums.ItemNotificationMessageType</a></td><td>Spécifie le type de message. Si le type a pour valeur <code>ProgressIndicator</code> ou <code>ErrorMessage</code>, une icône apparaît automatiquement et le message n’est pas permanent. Par conséquent, l’icône et les propriétés permanentes ne sont pas valides pour ces types de messages. Le fait de les inclure génère une exception <code>ArgumentException</code>. Si le type a pour valeur <code>ProgressIndicator</code>, le développeur doit supprimer ou remplacer l’indicateur de progression à la fin de l’action.</td></tr><tr><td><code>icon</code></td><td>String</td><td>Référence à une icône définie dans le manifeste dans la section <code>Resource</code>. Elle apparaît dans la barre d’informations. S’applique uniquement si le type a pour valeur <code>InformationalMessage</code>. Le fait de spécifier ce paramètre pour un type non pris en charge génère une exception.</td></tr><tr><td><code>message</code></td><td>String</td><td>Texte du message de notification. La longueur maximale est de 150 caractères. Si le développeur génère une chaîne plus longue, une exception <code>ArgumentOutOfRange</code> se déclenche.</td></tr><tr><td><code>persistent</code></td><td>Boolean</td><td>S’applique uniquement lorsque le type a pour valeur <code>InformationalMessage</code>. Sur <code>true</code>, le message est conservé jusqu’à ce qu’il soit supprimé par le complément ou masqué par l’utilisateur. Sur <code>false</code>, il est supprimé lorsque l’utilisateur accède à un autre élément. Pour les notifications d’erreur, le message est conservé jusqu’à ce qu’il soit vu par l’utilisateur. Le fait de spécifier ce paramètre pour un type non pris en charge génère une exception.</td></tr></tbody></table>|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult). |

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.3|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### Exemple

```
// Create three notifications, each with a different key
Office.context.mailbox.item.notificationMessages.addAsync("progress", {
  type: "progressIndicator",
  message : "An add-in is processing this message."
});
Office.context.mailbox.item.notificationMessages.addAsync("information", {
  type: "informationalMessage",
  message : "The add-in processed this message.",
  icon : "iconid",
  persistent: false
});
Office.context.mailbox.item.notificationMessages.addAsync("error", {
  type: "errorMessage",
  message : "The add-in failed to process this message."
});
```

####  getAllAsync([options], [callback])

Renvoie l’ensemble des clés et messages pour un élément.

##### Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult).

Une fois son exécution réussie, la propriété `asyncResult.value` contient un tableau des objets [`NotificationMessageDetails`](simple-types.md#notificationmessagedetails).|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.3|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### Exemple

```
// Get all notifications
Office.context.mailbox.item.notificationMessages.getAllAsync(function (asyncResult) {
  if (asyncResult.status != "failed") {
    Office.context.mailbox.item.notificationMessages.replaceAsync( "notifications", {
      type: "informationalMessage",
      message : "Found " + asyncResult.value.length + " notifications.",
      icon : "iconid",
      persistent: false
    });
  }
});
```

####  removeAsync(key, [options], [callback])

Supprime un message de notification pour un élément.

##### Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`key`| Chaîne||Clé pour le message de notification à supprimer.|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult).

Si la clé est introuvable, une erreur `KeyNotFound` est renvoyée dans la propriété `asyncResult.error`.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.3|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### Exemple

```
// Remove a notification
Office.context.mailbox.item.notificationMessages.removeAsync("progress");
```

####  replaceAsync(key, JSONmessage, [options], [callback])

Remplace un message de notification disposant d’une clé donnée par un autre message.

Si un message de notification avec la clé spécifiée n’existe pas, `replaceAsync` ajoute la notification.

##### Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`key`| String||Clé pour le message de notification à remplacer. Elle peut contenir jusqu’à 32 caractères.|
|`JSONmessage`| Objet||Objet JSON qui contient le nouveau message de notification qui va remplacer le message existant. Il se compose des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Description</th></tr></thead><tbody><tr><td><code>type</code></td><td><a href="Office.MailboxEnums.md#.ItemNotificationMessageType">Office.MailboxEnums.ItemNotificationMessageType</a></td><td>Spécifie le type de message. Si le type a pour valeur <code>ProgressIndicator</code> ou <code>ErrorMessage</code>, une icône apparaît automatiquement et le message n’est pas permanent. Par conséquent, l’icône et les propriétés permanentes ne sont pas valides pour ces types de messages. Le fait de les inclure génère une exception <code>ArgumentException</code>. Si le type a pour valeur <code>ProgressIndicator</code>, le développeur doit supprimer ou remplacer l’indicateur de progression à la fin de l’action.</td></tr><tr><td><code>icon</code></td><td>String</td><td>Référence à une icône définie dans le manifeste dans la section <code>Resource</code>. Elle apparaît dans la barre d’informations. S’applique uniquement si le type a pour valeur <code>InformationalMessage</code>. Le fait de spécifier ce paramètre pour un type non pris en charge génère une exception.</td></tr><tr><td><code>message</code></td><td>String</td><td>Texte du message de notification. La longueur maximale est de 150 caractères. Si le développeur génère une chaîne plus longue, une exception <code>ArgumentOutOfRange</code> se déclenche.</td></tr><tr><td><code>persistent</code></td><td>Boolean</td><td>S’applique uniquement lorsque le type a pour valeur <code>InformationalMessage</code>. Sur <code>true</code>, le message est conservé jusqu’à ce qu’il soit supprimé par le complément ou masqué par l’utilisateur. Sur <code>false</code>, il est supprimé lorsque l’utilisateur accède à un autre élément. Pour les notifications d’erreur, le message est conservé jusqu’à ce qu’il soit vu par l’utilisateur. Le fait de spécifier ce paramètre pour un type non pris en charge génère une exception.</td></tr></tbody></table>|
|`options`| Objet| &lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.<br/><br/>**Propriétés**<br/><table class="nested-table"><thead><tr><th>Nom</th><th>Type</th><th>Attributs</th><th>Description</th></tr></thead><tbody><tr><td><code>asyncContext</code></td><td>Objet</td><td>&lt;optional&gt;</td><td>Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</td></tr></tbody></table>|
|`callback`| fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult). |

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.3|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|

##### Exemple

```
// Replace a notification with an informational notification
Office.context.mailbox.item.notificationMessages.replaceAsync("progress", {
  type: "informationalMessage",
  message : "The message was processed successfully.",
  icon : "iconid",
  persistent: false
});
```
