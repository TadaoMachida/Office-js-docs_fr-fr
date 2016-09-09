
# Obtenir, définir ou ajouter des destinataires lors de la composition d’un rendez-vous ou d’un message dans Outlook


L’interface API JavaScript pour Office fournit des méthodes asynchrones ([Recipients.getAsync](../../reference/outlook/Recipients.md), [Recipients.setAsync](../../reference/outlook/Recipients.md) ou [Recipients.addAysnc](../../reference/outlook/Recipients.md)) pour respectivement obtenir, définir ou ajouter des destinataires dans le formulaire de composition d’un rendez-vous ou d’un message. Ces méthodes asynchrones sont disponibles uniquement pour les compléments de composition. Pour utiliser ces méthodes, assurez-vous que vous avez correctement configuré le manifeste du complément pour Outlook afin d’activer le complément dans les formulaires de composition, comme décrit dans la rubrique [Créer des compléments Outlook pour les formulaires de composition](../outlook/compose-scenario.md).

Certaines des propriétés qui représentent les destinataires dans un rendez-vous ou un message sont disponibles pour l’accès en lecture dans un formulaire de composition et de lecture. Ces propriétés sont [optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md) et [requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md) pour les rendez-vous et [cc](../../reference/outlook/Office.context.mailbox.item.md) et [to](../../reference/outlook/Office.context.mailbox.item.md) pour les messages. Dans un formulaire de lecture, vous pouvez accéder à la propriété directement dans l’objet parent, comme dans l’exemple suivant :




```js
item.cc
```

Toutefois, étant donné que l’utilisateur et votre complément peuvent insérer ou modifier un destinataire au même moment, vous devez, dans un formulaire de composition, utiliser la méthode asynchrone  **getAsync** pour obtenir ces propriétés, comme dans l’exemple suivant :




```js
item.cc.getAsync
```

Ces propriétés sont disponibles pour l’accès en écriture uniquement dans les formulaires de composition, pas dans les formulaires de lecture.

Comme avec la plupart des méthodes asynchrones dans l’interface API JavaScript pour Office, **getAsync**, **setAsync** et **addAsync** admettent des paramètres d’entrée facultatifs. Pour plus d’informations sur la spécification de ces paramètres d’entrée facultatifs, voir [Passage de paramètres facultatifs à des méthodes asynchrones](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) dans [Programmation asynchrone dans des compléments Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md).


## Pour obtenir les destinataires


Cette section présente un exemple de code qui obtient les destinataires d’un rendez-vous ou d’un message dont la composition est en cours et affiche les adresses de messagerie des destinataires. L’exemple de code suppose l’existence d’une règle dans le manifeste du complément qui active le complément dans un formulaire de composition pour un rendez-vous ou un message, comme indiqué ci-dessous. 


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

Dans l’interface API JavaScript pour Office, étant donné que les propriétés qui représentent les destinataires d’un rendez-vous ( **optionalAttendees** et **requiredAttendees**) sont différentes de celles d’un message ([bcc](../../reference/outlook/Office.context.mailbox.item.md),  **cc** et **to**), vous devez d’abord utiliser la propriété [item.itemType](../../reference/outlook/Office.context.mailbox.item.md) pour déterminer si l’élément dont la composition est en cours est un rendez-vous ou un message. En mode composition, toutes ces propriétés de rendez-vous et de messages sont des objets [Recipients](../../reference/outlook/Recipients.md), de sorte que vous pouvez ensuite appliquer la méthode asynchrone  **Recipients.getAsync**, pour obtenir les destinataires correspondants. 

Pour utiliser  **getAsync**, indiquez une méthode de rappel pour vérifier l’état, les résultats et les erreurs renvoyés par l’appel asynchrone  **getAsync**. Vous pouvez fournir des arguments à la méthode de rappel à l’aide du paramètre facultatif  _asyncContext_. La méthode de rappel renvoie un paramètre de sortie  _asyncResult_. Vous pouvez utiliser les propriétés  **status** et **error** de l’objet de paramètre [AsyncResult](../../reference/outlook/simple-types.md) pour vérifier l’état et les messages d’erreur de l’appel asynchrone, ainsi que la propriété **value** pour obtenir les destinataires réels. Les destinataires sont représentés dans un tableau d’objets [EmailAddressDetails](../../reference/outlook/simple-types.md).

Étant donné que la méthode  **getAsync** est asynchrone, si des actions ultérieures dépendent de l’obtention des destinataires, vous devez organiser votre code afin de ne lancer ces actions que dans la méthode de rappel correspondante, une fois que l’appel asynchrone a abouti.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get all the recipients of the composed item.
        getAllRecipients();
    });
}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous 
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients. 
            write ('To-recipients of the item:');
            displayAddresses(asyncResult);
        }    
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            write ('Cc-recipients of the item:');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get bcc-recipients of the item completed.
            // Display the email addresses of the bcc-recipients.
            write ('Bcc-recipients of the item:');
            displayAddresses(asyncResult);
        }
                        
        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
    for (var i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Pour définir des destinataires


Cette section présente un exemple de code qui définit les destinataires du rendez-vous ou du message que l’utilisateur compose. Le fait de définir des destinataires remplace tous les destinataires existants. Comme dans l’exemple précédent relatif à l’obtention des destinataires dans un formulaire de composition, cet exemple suppose que le complément est activé dans les formulaires de composition pour les rendez-vous et les messages. Cet exemple détermine d’abord si l’élément composé est un rendez-vous ou un message afin d’appliquer la méthode asynchrone  **Recipients.setAsync** sur les propriétés appropriées représentant les destinataires du rendez-vous ou du message.

Lorsque vous appelez  **setAsync**, fournissez un tableau comme argument d’entrée pour le paramètre  _recipients_, dans l’un des formats suivants :


- Un tableau de chaînes représentant des adresses SMTP.
    
- Un tableau de dictionnaires, chacun contenant un nom d’affichage et une adresse de messagerie, comme indiqué dans l’exemple de code suivant.
    
- Un tableau d’objets  **EmailAddressDetails**, semblable à celui renvoyé par la méthode  **getAsync**.
    
Vous pouvez éventuellement fournir une méthode de rappel comme argument d’entrée pour la méthode  **setAsync** afin de vous assurer que tout code qui dépend de la définition des destinataires ne s’exécute que lorsque l’opération aboutit. Vous pouvez également fournir des arguments à la méthode de rappel à l’aide du paramètre facultatif _asyncContext_. Si vous utilisez une méthode de rappel, vous pouvez accéder à un paramètre d’entrée  _asyncResult_ et utiliser les propriétés **status** et **error** de l’objet de paramètre **AsyncResult** pour vérifier l’état et les messages d’erreur de l’appel asynchrone.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set recipients of the composed item.
        setRecipients();
    });
}

// Set the display name and email addresses of the recipients of 
// the composed item.
function setRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;

    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    toRecipients.setAsync(
        [{
            "displayName":"Graham Durkin", 
            "emailAddress":"graham@contoso.com"
         },
         {
            "displayName" : "Donnie Weinberg",
            "emailAddress" : "donnie@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
             "displayName":"Perry Horning", 
             "emailAddress":"perry@contoso.com"
         },
         {
             "displayName" : "Guy Montenegro",
             "emailAddress" : "guy@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.


    // If the item has the bcc field, i.e., item is message,
    // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                 "displayName":"Lewis Cate", 
                 "emailAddress":"lewis@contoso.com"
             },
             {
                 "displayName" : "Francisco Stitt",
                 "emailAddress" : "francisco@contoso.com"
             }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to set bcc-recipients of the item completed.
                    // Do whatever appropriate for your scenario.
                }
        }); // End bcc setAsync.
    }
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```


## Pour ajouter des destinataires


Si vous ne souhaitez pas remplacer les destinataires existants dans un rendez-vous ou un message, vous pouvez utiliser la méthode asynchrone  **Recipients.addAsync** à la place de **Recipients.setAsync** pour ajouter des destinataires. La méthode **addAsync** fonctionne de manière semblable à la méthode **setAsync** dans la mesure où elle requiert un argument d’entrée _recipients_. Vous pouvez éventuellement fournir une méthode de rappel et tous les arguments pour le rappel à l’aide du paramètre asyncContext. Vous pouvez vérifier l’état, le résultat et les erreurs de l’appel asynchrone  **addAsync** en utilisant le paramètre d’entrée _asyncResult_ de la méthode de rappel. L’exemple suivant vérifie que l’élément en cours de composition est un rendez-vous et y ajoute deux participants obligatoires.


```js
// Add specified recipients as required attendees of
// the composed appointment. 
function addAttendees() {
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName":"Kristie Jensen", 
            "emailAddress":"kristie@contoso.com"
         },
         {
            "displayName" : "Pansy Valenzuela",
            "emailAddress" : "pansy@contoso.com"
          }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to add attendees completed.
                // Do whatever appropriate for your scenario.
            }
        }); // End addAsync.
    }
}
```


## Ressources supplémentaires



- [Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Obtention et définition de données d’élément Outlook dans des formulaires de lecture ou de composition](../outlook/item-data.md)
    
- [Créer des compléments Outlook pour les formulaires de composition](../outlook/compose-scenario.md)
    
- [Programmation asynchrone dans des compléments Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [Obtenir ou définir l’objet lors de la composition d’un rendez-vous ou d’un message dans Outlook](../outlook/get-or-set-the-subject.md)
    
- [Insérer des données dans le corps lors de la composition d’un rendez-vous ou d’un message dans Outlook](../outlook/insert-data-in-the-body.md)
    
- [Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook](../outlook/get-or-set-the-location-of-an-appointment.md)
    
- [Obtenir ou définir l’heure lors de la composition d’un rendez-vous dans Outlook](../outlook/get-or-set-the-time-of-an-appointment.md)
    
