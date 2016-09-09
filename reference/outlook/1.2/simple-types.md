

# Types simples

####  AsyncResult

Objet qui encapsule le résultat d’une requête asynchrone, y compris les informations d’état et d’erreur quand la demande a échoué.

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`asyncContext`| Objet|Obtient l’objet transmis au paramètre facultatif `asyncContext` de la méthode appelée dans le même état que celui dans lequel il a été transmis.|
|`error`| Erreur|Obtient un objet Error qui fournit une description de l’erreur, si une erreur s’est produite.|
|`status`| [Office.AsyncResultStatus](Office.md#asyncresultstatus-string)|Obtient l’état de l’opération asynchrone.|
|`value`| Objet|Obtient la charge utile ou le contenu de l’opération asynchrone, le cas échéant.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|Mode Outlook applicable| Composition ou lecture|
#### AttachmentDetails

Représente la pièce jointe d’un élément du serveur. Mode lecture uniquement.

Un tableau d’objets `AttachmentDetail` est renvoyé comme propriété `attachments` d’un objet `Appointment` ou `Message`.

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`attachmentType`| [Office.MailboxEnums.AttachmentType](Office.MailboxEnums.md#attachmenttype-string)|Obtient une valeur qui indique le type d’une pièce jointe.|
|`contentType`| Chaîne|Obtient le type de contenu MIME de la pièce jointe.|
|`id`| String|Obtient l’ID de pièce jointe Exchange de la pièce jointe.|
|`isInline`| Boolean|Obtient une valeur indiquant si la pièce jointe doit être affichée dans le corps de l’élément.|
|`name`| String|Obtient le nom de la pièce jointe.|
|`size`| Nombre|Obtient la taille de la pièce jointe en octets.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|
#### Contact

Représente un contact stocké sur le serveur. Mode lecture uniquement.

La liste des contacts associés à un message électronique ou un rendez-vous est renvoyée dans la propriété `contacts` de l’objet [`Entities`](simple-types.md#entities), qui est renvoyé par la méthode `getEntities` ou `getEntitiesByType` de l’élément actif.

##### Propriétés :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`addresses`| Array.&lt;String&gt;| &lt;nullable&gt;|Tableau de chaînes contenant les adresses de messagerie et postales associées au contact.|
|`businessName`| Chaîne| &lt;nullable&gt;|Chaîne contenant le nom de l’entreprise associée au contact.|
|`emailAddresses`| Array.&lt;String&gt;| &lt;nullable&gt;|Tableau de chaînes contenant les adresses de messagerie SMTP associées au contact.|
|`personName`| String| &lt;nullable&gt;|Chaîne contenant le nom de la personne associée au contact.|
|`phoneNumbers`| Array.&lt;[PhoneNumber](simple-types.md#phonenumber)&gt;| &lt;nullable&gt;|Tableau contenant un objet `PhoneNumber` pour chaque numéro de téléphone associé au contact.|
|`urls`| Array.&lt;String&gt;| &lt;nullable&gt;|Tableau de chaînes contenant les URL Internet associées au contact.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| Restricted|
|Mode Outlook applicable| Lecture|
####  EmailAddressDetails

Fournit les propriétés relatives à l’expéditeur ou aux destinataires spécifiés d’un e-mail ou d’un rendez-vous.

##### Type :

*   Objet

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`appointmentResponse`| [Office.MailboxEnums.ResponseType](Office.MailboxEnums.md#responsetype-string)|Obtient la réponse d’un participant pour un rendez-vous. Cette propriété s’applique uniquement aux participants d’un rendez-vous, tel que représenté par la propriété [`optionalAttendees`](Office.context.mailbox.item.md#optionalattendees-arrayemailaddressdetailsrecipients) ou [`requiredAttendees`](Office.context.mailbox.item.md#requiredattendees-arrayemailaddressdetailsrecipients). Cette propriété renvoie `undefined` dans d’autres scénarios.|
|`displayName`| String|Obtient le nom d’affichage associé à une adresse de messagerie.|
|`emailAddress`| String|Obtient l’adresse de messagerie SMTP.|
|`recipientType`| [Office.MailboxEnums.RecipientType](Office.MailboxEnums.md#recipienttype-string)|Obtient le type d’adresse de messagerie d’un destinataire.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Composition ou lecture|
#### EmailUser

Représente un compte de messagerie sur un serveur Exchange.

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`displayName`| Chaîne|Obtient le nom d’affichage associé à une adresse de messagerie.|
|`emailAddress`| String|Obtient l’adresse de messagerie SMTP.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|
#### Entités

Représente une collection d’entités trouvées dans un message électronique ou un rendez-vous. Mode lecture uniquement.

L’objet `Entities` est un conteneur pour les tableaux d’entités renvoyés par les méthodes `getEntities` et `getEntitiesByType` quand l’élément (message électronique ou rendez-vous) contient une ou plusieurs entités qui ont été trouvées par le serveur. Vous pouvez utiliser ces entités dans votre code pour fournir des informations de contexte supplémentaires, par exemple une carte montrant une adresse trouvée dans l’élément, ou pour ouvrir un numéroteur quand un numéro de téléphone est trouvé dans l’élément.

Si aucune entité du type spécifié dans la propriété n’est présente dans l’élément, la propriété associée à cette entité a la valeur `null`. Par exemple, si un message contient une adresse et un numéro de téléphone, les propriétés `addresses` et `phoneNumbers` contiennent des informations, alors que les autres propriétés ont la valeur `null`.

Pour être reconnue en tant qu’adresse, la chaîne doit contenir une adresse postale incluant au moins un sous-ensemble d’éléments tels que le numéro de rue, le nom de rue, la ville, le département/la région/l’État et le code postal.

Pour être reconnue comme numéro de téléphone, la chaîne doit contenir un numéro de téléphone de type nord-américain.

La reconnaissance d’entité repose sur la reconnaissance du langage naturel qui est basée sur l’apprentissage par l’ordinateur de grandes quantités de données. La reconnaissance d’une entité n’est pas déterministe et sa réussite s’appuie parfois sur le contexte particulier de l’élément.

Quand les tableaux de propriétés sont renvoyés par la méthode `getEntitiesByType`, seule la propriété de l’entité spécifiée contient des données ; toutes les autres propriétés ont la valeur `null`.

##### Propriétés :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`addresses`| Array.&lt;String&gt;| &lt;nullable&gt;|Obtient les adresses physiques (rue ou adresse postale) trouvées dans un e-mail ou un rendez-vous.|
|`contacts`| Array.&lt;[Contact](simple-types.md#contact)&gt;| &lt;nullable&gt;|Obtient les contacts trouvés dans une adresse de messagerie ou un rendez-vous.|
|`emailAddresses`| Array.&lt;String&gt;| &lt;nullable&gt;|Obtient les adresses de messagerie trouvées dans un e-mail ou un rendez-vous.|
|`meetingSuggestions`| Array.&lt;[MeetingSuggestion](simple-types.md#meetingsuggestion)&gt;| &lt;nullable&gt;|Obtient les suggestions de réunion trouvées dans un e-mail.|
|`phoneNumbers`| Array.&lt;[PhoneNumber](simple-types.md#phonenumber)&gt;| &lt;nullable&gt;|Obtient les numéros de téléphone trouvés dans un e-mail ou un rendez-vous.|
|`taskSuggestions`| Array.&lt;[TaskSuggestion](simple-types.md#tasksuggestion)&gt;| &lt;nullable&gt;|Obtient les suggestions de tâche trouvées dans un e-mail ou un rendez-vous.|
|`urls`| Array.&lt;String&gt;| &lt;nullable&gt;|Obtient les URL Internet présentes dans un e-mail ou un rendez-vous.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|
#### LocalClientTime

Représente une date et une heure dans le fuseau horaire du client. Mode lecture uniquement.

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`month`| Nombre|Nombre entier représentant le mois (0 correspondant au mois de janvier et 11 à décembre).|
|`date`| Nombre|Nombre entier représentant le jour du mois.|
|`year`| Nombre|Nombre entier représentant l’année.|
|`hours`| Nombre|Nombre entier représentant l’heure sur une horloge de 24 heures.|
|`minutes`| Nombre|Nombre entier représentant le nombre de minutes.|
|`seconds`| Nombre|Nombre entier représentant le nombre de secondes.|
|`milliseconds`| Nombre|Nombre entier représentant le nombre de millisecondes.|
|`timezoneOffset`| Nombre|Nombre entier représentant l’écart en minutes entre le fuseau horaire local et l’heure UTC.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|
#### MeetingSuggestion

Représente une réunion proposée trouvée dans un élément. Mode lecture uniquement.

La liste des réunions proposées dans un message électronique est renvoyée dans la propriété `meetingSuggestions` de l’objet [`Entities`](simple-types.md#entities) renvoyé lorsque la méthode [`getEntities`](Office.context.mailbox.item.md#getentities--entities) ou [`getEntitiesByType`](Office.context.mailbox.item.md#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) est appelée sur l’élément actif.

Les valeurs `start` et `end` sont des représentations sous forme de chaîne d’un objet Date qui contient la date et l’heure à laquelle la réunion proposée doit commencer et se terminer. Les valeurs sont comprises dans le fuseau horaire par défaut spécifié pour l’utilisateur actif.

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`attendees`| Array.&lt;[EmailUser](simple-types.md#emailuser)&gt;|Obtient les participants à une réunion suggérée.|
|`end`| String|Obtient la date et l’heure de fin d’une réunion suggérée.|
|`location`| Chaîne|Obtient le lieu d’une réunion suggérée.|
|`meetingString`| String|Obtient une chaîne identifiée en tant que suggestion de réunion.|
|`start`| Chaîne|Obtient la date et l’heure de début d’une réunion suggérée.|
|`subject`| String|Obtient l’objet d’une réunion suggérée.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|
#### PhoneNumber

Représente un numéro de téléphone identifié dans un élément. Mode lecture uniquement.

Un tableau d’objets `PhoneNumber` contenant des numéros de téléphone trouvés dans un message électronique est renvoyé dans la propriété `phoneNumbers` de l’objet [`Entities`](simple-types.md#entities) renvoyé lors de l’appel de la méthode [`getEntities`](Office.context.mailbox.item.md#getentities--entities) pour l’élément sélectionné.

##### Type :

*   Objet

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`originalPhoneString`| String|Obtient le texte identifié dans un élément en tant que numéro de téléphone.|
|`phoneString`| String|Obtient une chaîne contenant un numéro de téléphone. Cette chaîne contient uniquement les chiffres du numéro de téléphone et exclut les caractères tels que les parenthèses et les caractères, s’ils existent dans l’élément d’origine.|
|`type`| String|Obtient une chaîne qui identifie le type de numéro de téléphone : `Home`, `Work`, `Mobile`, `Unspecified`.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|
#### TaskSuggestion

Représente une suggestion de tâche identifiée dans un élément. Mode lecture uniquement.

La liste des tâches proposées dans un message électronique est renvoyée dans la propriété `taskSuggestions` de l’objet [`Entities`Entities`Entities`](simple-types.md#entities) renvoyé lorsque la méthode [`getEntities`](Office.context.mailbox.item.md#getentities--entities) ou [`getEntitiesByType`](Office.context.mailbox.item.md#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) est appelée sur l’élément actif.

##### Propriétés :

|Nom| Type| Description|
|---|---|---|
|`assignees`| Array.&lt;[EmailUser](simple-types.md#emailuser)&gt;|Obtient les utilisateurs auxquels une tâche suggérée doit être affectée.|
|`taskString`| String|Obtient le texte d’un élément identifié en tant que suggestion de tâche.|

##### Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../tutorial-api-requirement-sets.md)| 1,0|
|[Niveau d’autorisation minimal](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Mode Outlook applicable| Lecture|
