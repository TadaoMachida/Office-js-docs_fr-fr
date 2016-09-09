 

# MailboxEnums

## [Office](Office.md). MailboxEnums

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition ou lecture|

### Membres

#### AttachmentType :String

Spécifie le type d’une pièce jointe.

AttachmentType

##### Type :

*   String

##### Propriétés :

|Nom| Type| Valeur | Description|
|---|---|---|---|
|`File`| String|`file`|La pièce jointe est un fichier.|
|`Item`| String|`item`|La pièce jointe est un élément Exchange.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition ou lecture|
#### EntityType :String

Spécifie le type d’une entité.

EntityType

##### Type :

*   String

##### Propriétés :

|Nom| Type| Valeur | Description|
|---|---|---|---|
|`Address`| String|`address`|Spécifie que l’entité est une adresse postale.|
|`Contact`| String|`contact`|Spécifie que l’entité est un contact.|
|`EmailAddress`| String|`emailAddress`|Spécifie que l’entité est une adresse de messagerie SMTP.|
|`MeetingSuggestion`| String|`meetingSuggestion`|Spécifie que l’entité est une suggestion de réunion.|
|`PhoneNumber`| String|`phoneNumber`|Spécifie que l’entité est un numéro de téléphone.|
|`TaskSuggestion`| Chaîne|`taskSuggestion`|Spécifie que l’entité est une suggestion de tâche.|
|`URL`| String|`url`|Spécifie que l’entité est une URL Internet.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition ou lecture|
#### ItemNotificationMessageType :String

Spécifie le type de message de notification pour un rendez-vous ou un message.

ItemNotificationMessageType

##### Type :

*   String

##### Propriétés :

|Nom| Type| Valeur | Description|
|---|---|---|---|
|`ProgressIndicator`| String|`progressIndicator`|Le message de notification est un indicateur de progression.|
|`InformationalMessage`| String|`informationalMessage`|Le message de notification est un message d’information.|
|`ErrorMessage`| String|`errorMessage`|Le message de notification est un message d’erreur.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.3|
|Mode Outlook applicable| Composition ou lecture|
#### ItemType :String

Spécifie le type d’un élément.

ItemType

##### Type :

*   String

##### Propriétés :

|Nom| Type| Valeur | Description|
|---|---|---|---|
|`Message`| Chaîne|`message`|Message électronique, demande de réunion, réponse à une demande de réunion ou annulation d’une réunion.|
|`Appointment`| Chaîne|`appointment`|Élément de rendez-vous.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition ou lecture|
#### RecipientType :String

Spécifie le type de destinataire d’un rendez-vous.

RecipientType

##### Type :

*   String

##### Propriétés :

|Nom| Type| Valeur | Description|
|---|---|---|---|
|`Other`| String|`other`|Le destinataire ne fait pas partie des autres types de destinataires.|
|`DistributionList`| String|`distributionList`|Le destinataire est une liste de distribution contenant une liste d’adresses de messagerie.|
|`User`| Chaîne|`user`|Le destinataire est une adresse de messagerie SMTP qui se trouve sur le serveur Exchange.|
|`ExternalUser`| String|`externalUser`|Le destinataire est une adresse de messagerie SMTP qui ne se trouve pas sur le serveur Exchange.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.1|
|Mode Outlook applicable| Composition ou lecture|
#### ResponseType :String

Spécifie le type de réponse à une invitation à une réunion.

ResponseType

##### Type :

*   String

##### Propriétés :

|Nom| Type| Valeur | Description|
|---|---|---|---|
|`None`| Chaîne|`none`|Il n’y a eu aucune réponse du participant.|
|`Organizer`| String|`organizer`|Le participant est l’organisateur de la réunion.|
|`Tentative`| String|`tentative`|La demande de réunion a été provisoirement acceptée par le participant.|
|`Accepted`| String|`accepted`|La demande de réunion a été acceptée par le participant.|
|`Declined`| String|`declined`|La demande de réunion a été refusée par le participant.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition ou lecture|

#### RestVersion :String

Spécifie la version de l’API REST qui correspond à un ID d’élément au format REST. 

RestVersion

##### Type :

*   String

##### Propriétés :

|Nom| Type| Valeur | Description|
|---|---|---|---|
|`v1_0`| String|`v1.0`|Version 1.0.|
|`v2_0`| String|`v2.0`|Version 2.0.|
|`Beta`| String|`beta`|Bêta.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](./tutorial-api-requirement-sets.md)| 1.3|
|Mode Outlook applicable| Composition ou lecture|
