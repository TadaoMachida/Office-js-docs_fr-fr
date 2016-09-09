 

# MailboxEnums

## [Office](Office.md). MailboxEnums

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition|

### Membres

#### AttachmentType :String

Spécifie le type d’une pièce jointe. Mode composition uniquement.

AttachmentType

##### Type :

*   String

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`File`| String|La pièce jointe est un fichier.|
|`Item`| String|La pièce jointe est un élément Exchange.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition|
#### EntityType :String

Spécifie le type d’une entité. Mode composition uniquement.

EntityType

##### Type :

*   String

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Address`| String|Spécifie que l’entité est une adresse postale.|
|`Contact`| String|Spécifie que l’entité est un contact.|
|`EmailAddress`| String|Spécifie que l’entité est une adresse de messagerie SMTP.|
|`MeetingSuggestion`| String|Spécifie que l’entité est une suggestion de réunion.|
|`PhoneNumber`| String|Spécifie que l’entité est un numéro de téléphone.|
|`TaskSuggestion`| Chaîne|Spécifie que l’entité est une suggestion de tâche.|
|`URL`| String|Spécifie que l’entité est une URL Internet.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition|
#### ItemType :String

Spécifie le type d’un élément. Mode composition uniquement.

ItemType

##### Type :

*   String

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Message`| Chaîne|Message électronique, demande de réunion, réponse à une demande de réunion ou annulation d’une réunion.|
|`Appoinment`| Chaîne|Élément de rendez-vous.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition|
#### RecipientType :String

Spécifie le type de destinataire d’un rendez-vous. Mode composition uniquement.

RecipientType

##### Type :

*   String

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Other`| String|Le destinataire ne fait pas partie des autres types de destinataires.|
|`DistributionList`| String|Le destinataire est une liste de distribution contenant une liste d’adresses de messagerie.|
|`User`| Chaîne|Le destinataire est une adresse de messagerie SMTP qui se trouve sur le serveur Exchange.|
|`ExternalUser`| String|Le destinataire est une adresse de messagerie SMTP qui ne se trouve pas sur le serveur Exchange.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1.1|
|Mode Outlook applicable| Composition|
#### ResponseType :String

Spécifie le type de réponse à une invitation à une réunion. Mode composition uniquement.

ResponseType

##### Type :

*   String

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`None`| Chaîne|Il n’y a eu aucune réponse du participant.|
|`Organizer`| String|Le participant est l’organisateur de la réunion.|
|`Tentative`| String|La demande de réunion a été provisoirement acceptée par le participant.|
|`Accepted`| String|La demande de réunion a été acceptée par le participant.|
|`Declined`| String|La demande de réunion a été refusée par le participant.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition|
