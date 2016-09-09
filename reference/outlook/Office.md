 

# Bureau

L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API partagée](../shared/shared-api.md).

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition ou lecture|

### Espaces de noms

[context](Office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.

[MailboxEnums](Office.MailboxEnums.md) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.

### Membres

####  AsyncResultStatus :String

Spécifie le résultat d’un appel asynchrone.

##### Type :

*   String

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Succeeded`| String|L’appel a réussi.|
|`Failed`| Chaîne|L’appel n’a pas réussi.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition ou lecture|
####  CoercionType :String

Indique comment forcer le type des données retournées ou définies par la méthode appelée.

##### Type :

*   String

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Html`| Chaîne|Demande que les données soient renvoyées au format HTML.|
|`Text`| Chaîne|Demande que les données soient renvoyées au format texte.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition ou lecture|
####  SourceProperty :String

Spécifie la source des données renvoyées par la méthode appelée.

##### Type :

*   String

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Body`| Chaîne|La source de données est dans le corps d’un message.|
|`Subject`| String|La source de données est dans l’objet d’un message.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition ou lecture|
