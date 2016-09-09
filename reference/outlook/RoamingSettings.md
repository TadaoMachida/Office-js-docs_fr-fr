

# RoamingSettings

Les paramètres créés à l’aide des méthodes de l’objet `RoamingSettings` sont enregistrés par complément et par utilisateur. En d’autres termes, ils ne sont disponibles que pour le complément qui les a créés et uniquement dans la boîte aux lettres de l’utilisateur où ils sont enregistrés.

> Même si l’API du complément Outlook limite l’accès à ces paramètres au complément qui les a créés, ces paramètres ne doivent pas être considérés comme un espace de stockage sécurisé. Ils sont accessibles via les services web Exchange ou l’interface MAPI étendue. Nous vous recommandons de ne pas les utiliser pour stocker des informations sensibles telles que des informations d’identification ou des jetons de sécurité.

Le nom d’un paramètre est une donnée String, alors que sa valeur peut être une donnée String, Number, Boolean, Null, Object ou Array.

L’objet `RoamingSettings` est accessible via la propriété [`roamingSettings`](Office.context.md#roamingsettings-roamingsettings) dans l’espace de noms `Office.context`.

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Mode Outlook applicable| Composition ou lecture|

### Exemple

```
// Get the current value of the 'myKey' setting
var value = Office.context.roamingSettings.get('myKey');
// Update the value of the 'myKey' setting
Office.context.roamingSettings.set('myKey', 'Hello World!');
// Persist the change
Office.context.roamingSettings.saveAsync();
```

### Méthodes

####  get(name) → (nullable) {String|Number|Boolean|Object|Array}

Récupère le paramètre spécifié.

##### Paramètres :

|Nom| Type| Description|
|---|---|---|
|`name`| String|Nom respectant l’emploi des majuscules et minuscules pour le paramètre à récupérer.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Mode Outlook applicable| Composition ou lecture|

##### Renvoie :

<dl class="param-type">

<dt>Type</dt>

<dd>String | Number | Boolean | Object | Array</dd>

</dl>

####  remove(name)

Supprime le paramètre spécifié.

##### Paramètres :

|Nom| Type| Description|
|---|---|---|
|`name`| String|Nom respectant l’emploi des majuscules et minuscules pour le paramètre à supprimer.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Mode Outlook applicable| Composition ou lecture|
####  saveAsync([callback])

Enregistre les paramètres.

Tous les paramètres précédemment enregistrés par un complément sont chargés lorsqu’il est initialisé. Ainsi, pendant la durée de la session, il vous suffit d’employer les méthodes [`set`](RoamingSettings.md#setname-value) et [`get`](RoamingSettings.md#getname--nullable-stringnumberbooleanobjectarray) pour utiliser la copie en mémoire du conteneur de propriétés de paramètres. Pour conserver les paramètres et faire en sorte qu’ils soient disponibles lors de la prochaine utilisation du complément, utilisez la méthode `saveAsync`.

##### Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`callback`| function| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](simple-types.md#asyncresult). |

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Mode Outlook applicable| Composition ou lecture|
####  set(name, value)

Définit ou crée le paramètre spécifié.

La méthode set permet de créer un paramètre du nom spécifié s’il n’existe pas déjà, ou de définir un paramètre existant du nom spécifié. La valeur est stockée dans le document sous forme de représentation JSON sérialisée de son type de données.

Un maximum de 2 Mo est disponible pour les paramètres de chaque complément, et chaque paramètre est limité à 32 Ko.

Les modifications apportées aux paramètres à l’aide de la fonction `set` ne sont pas enregistrées sur le serveur tant que la fonction [`saveAsync`](RoamingSettings.md#saveasynccallback) n’est pas appelée.

##### Paramètres :

|Nom| Type| Description|
|---|---|---|
|`name`| String|Nom qui respecte la casse du paramètre à définir ou créer.|
|`value`| String &#124; Number &#124; Boolean &#124; Object &#124; Array|Valeur à stocker.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Mode Outlook applicable| Composition ou lecture|
