
# Ajouter et supprimer des pièces jointes à un élément dans un formulaire de composition dans Outlook

Vous pouvez utiliser les méthodes [addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) et [addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md) pour joindre respectivement un fichier et un élément Outlook à l’élément en cours de composition par l’utilisateur. Les deux méthodes sont asynchrones, ce qui signifie que l’exécution peut se poursuivre sans attendre que l’action d’ajout de pièce jointe se termine. Selon l’emplacement d’origine et la taille de la pièce jointe en cours d’ajout, l’exécution de l’appel asynchrone d’ajout de pièce jointe peut prendre un certain temps. Si des tâches dépendent de l’exécution de l’action, vous devez les réaliser dans une méthode de rappel. Cette méthode de rappel est facultative et elle est appelée lorsque le téléchargement de la pièce jointe est terminé. La méthode de rappel admet un objet [AsyncResult](http://dev.outlook.com/reference/add-ins/simple-types.md) comme paramètre de sortie fournissant un état, une erreur et une valeur renvoyée à partir de l’action d’ajout de pièce jointe. Si le rappel exige des paramètres supplémentaires, vous pouvez les spécifier dans le paramètre facultatif _options.aysncContext_.  _options.asyncContext_ peut être de n’importe quel type attendu par votre méthode de rappel.

Par exemple, vous pouvez définir _options.asyncContext_ comme objet JSON qui contient au moins une paire clé-valeur, avec le caractère « : » séparant une clé et la valeur, et un caractère « , » séparant une paire clé-valeur d’une autre. Vous pouvez trouver plus d’exemples sur le [passage de paramètres facultatifs à des méthodes asynchrones](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) dans la plateforme des Compléments Office dans [Programmation asynchrone dans des compléments Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md). L’exemple suivant montre comment utiliser le paramètre **asyncContext** pour passer 2 arguments à une méthode de rappel :




```js
{ asyncContext: { var1: 1, var2: 2} }
```

Vous pouvez vérifier la réussite ou l’erreur d’un appel de méthode asynchrone dans la méthode de rappel à l’aide des propriétés  **status** et **error** de l’objet **AsyncResult**. Si l’association de la pièce jointe aboutit, vous pouvez utiliser la propriété  **AsyncResult.value** pour obtenir l’ID de la pièce jointe. Il s’agit d’un nombre entier que vous pouvez ensuite utiliser pour supprimer la pièce jointe.


 >**Remarque**  Il est recommandé d’utiliser l’ID de pièce jointe pour la supprimer uniquement si le même complément a ajouté cette pièce jointe dans la même session. Dans Outlook Web App et OWA pour périphériques, l’ID de pièce jointe est valide uniquement dans la même session. Une session est terminée lorsque l’utilisateur ferme le complément, ou si l’utilisateur commence la composition dans un formulaire incorporé, avant de fermer ce formulaire pour continuer dans une fenêtre distincte.


## Attachement d’un fichier

Vous pouvez joindre un fichier à un message ou un rendez-vous dans un formulaire de composition en utilisant la méthode  **addFileAttachmentAsync** et en spécifiant l’URI du fichier. Si le fichier est protégé, vous pouvez inclure une identité appropriée ou un jeton d’authentification comme paramètre de chaîne de requête d’URI. Exchange effectuera un appel à l’URI pour obtenir la pièce jointe, et le service web qui protège le fichier devra utiliser le jeton comme moyen d’authentification.

L’exemple JavaScript suivant est un complément de composition qui joint un fichier, picture.png, à partir d’un serveur web au message ou rendez-vous en cours de composition. La méthode de rappel prend  **asyncResult** comme paramètre, vérifie l’état de l’attachement, et obtient l’ID de pièce jointe si l’attachement aboutit.




```js
var mailbox;
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

Office.initialize = function () {
    mailbox = Office.context.mailbox;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback method is invoked and gets the attachment ID. 
        // You can optionally pass any object that you would  
        // access in the callback method as an argument to  
        // the asyncContext parameter.
        mailbox.item.addFileAttachmentAsync(
            attachmentURI,
            'picture.png',
            { asyncContext: null },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Get the ID of the attached file.
                    attachmentID = asyncResult.value;
                    write('ID of added attachment: ' + attachmentID);
                }
            });
    });
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Attachement d’un élément Outlook

Vous pouvez joindre un élément Outlook (par exemple, un élément de messagerie, de calendrier ou de contact) à un message ou à un rendez-vous dans un formulaire de composition en précisant l’ID des services web Exchange (EWS) de l’élément et en utilisant la méthode  **addItemAttachmentAsync**. Vous pouvez obtenir l’ID EWS d’un élément de messagerie, de calendrier, de contact ou de tâche dans la boîte aux lettres de l’utilisateur en utilisant la méthode [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) et en accédant à l’opération EWS [FindItem](http://msdn.microsoft.com/en-us/library/ebad6aae-16e7-44de-ae63-a95b24539729%28Office.15%29.aspx). La propriété [item.itemId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.md) fournit également l’ID EWS d’un élément existant dans un formulaire de lecture.

La fonction JavaScript suivante,  `addItemAttachment`, étend le premier exemple ci-dessus, et ajoute un élément comme pièce jointe à l’e-mail ou au rendez-vous en cours de composition. La fonction prend comme argument l’ID EWS de l’élément qui doit être joint. Si l’attachement aboutit, l’ID de pièce jointe est obtenu pour un traitement ultérieur, y compris la suppression de cette pièce jointe dans la même session.




```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(ID) {
    // When the attachment finishes uploading, the
    // callback method is invoked. Here, the callback
    // method uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to 
    // access in the callback method as an argument to 
    // the asyncContext parameter.
    mailbox.item.addItemAttachmentAsync(
        ID,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```


 >**Remarque**  Vous pouvez utiliser un complément de composition pour joindre une instance d’un rendez-vous périodique dans Outlook Web App ou OWA pour périphériques. Cependant, dans le client riche Outlook de prise en charge, la tentative d’attachement d’une instance entraîne l’attachement d’une série périodique (rendez-vous principal).


## Suppression d’une pièce jointe


Vous pouvez supprimer une pièce jointe de fichier ou d’élément d’un élément de rendez-vous ou de message dans un formulaire de composition en indiquant l’ID de pièce jointe correspondant et en utilisant la méthode [removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md). Vous devez supprimer uniquement les pièces jointes que le même complément a ajoutées dans la même session. Vous devez vous assurer que l’ID de pièce jointe correspond à une pièce jointe valide, sinon la méthode renverra une erreur. À l’instar des méthodes  **addFileAttachmentAsync** et **addItemAttachmentAsync**,  **removeAttachmentAsync** est une méthode asynchrone. Vous devez fournir une méthode de rappel pour vérifier l’état et toute erreur en utilisant l’objet de paramètre de sortie **AsyncResult**. Vous pouvez également passer des paramètres supplémentaires à la méthode de rappel à l’aide du paramètre facultatif  **asyncContext**, qui est un objet JSON de paires clé-valeur.

La fonction JavaScript suivante,  `removeAttachment`, continue d’étendre les exemples ci-dessus, et supprime la pièce jointe indiquée dans l’e-mail ou le rendez-vous en cours de composition. La fonction prend comme argument l’ID de la pièce jointe à supprimer. Vous pouvez obtenir l’ID d’une pièce jointe après un appel de la méthode  **addFileAttachmentAsync** ou **addItemAttachmentAsync** réussi, et le stocker pour un appel de la méthode **removeAttachmentAsync** ultérieur.




```js
// Removes the specified attachment from the composed item.
// ID is the Exchange identifier of the attachment to be 
// removed. 
function removeAttachment(ID) {
    // When the attachment is removed, the
    // callback method is invoked. Here, the callback
    // method uses an asyncResult parameter and gets
    // the ID of the removed attachment if the removal
    // succeeds.
    // You can optionally pass any object you wish to 
    // access in the callback method as an argument to 
    // the asyncContext parameter.
    mailbox.item.removeAttachmentAsync(
        ID,
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                write('Removed attachment with the ID: ' + asyncResult.value);
            }
        });
}
```


## Conseils en matière d’ajout et de suppression de pièces jointes


Si votre complément de composition ajoute et supprime des pièces jointes, structurez votre code de façon à passer un ID de pièce jointe valide à l’appel de suppression de pièce jointe et à gérer le cas de figure où  **AsyncResult.error** renvoie **InvalidAttachmentId**. Selon l’emplacement et la taille d’une pièce jointe, l’exécution de l’attachement d’un fichier ou d’un élément peut prendre un certain temps. L’exemple suivant contient un appel à  **addFileAttachmentAsync**,  `write` et **removeAttachmentAsync**. Vous pouvez penser que les appels s’exécutent de façon séquentielle l’un après l’autre.


```js
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

// Gets the current time in minutes, seconds and milliseconds.
function minutesSecondsMilliSeconds()
{
    var d = new Date();
    return d.getMinutes() + ":" + d.getSeconds() + ":" + d.getMilliseconds();
}

Office.context.mailbox.item.addFileAttachmentAsync(
        attachmentURI,
        'Welcome document',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(1): ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('(2): ' + minutesSecondsMilliSeconds() + ' ' + 
                    'ID of added attachment: ' + attachmentID);
            }
            write ('(3): ' + minutesSecondsMilliSeconds() + ' ' + 
                'Finishing addFileAttachmentAsync callback method.');
        });

write ('(4): ' + minutesSecondsMilliSeconds() + ' ' + 
    'attachmentID is: ' + attachmentID);

Office.context.mailbox.item.removeAttachmentAsync(
        attachmentID,      
        { asyncContext: null },
       function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(5): ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {           
                write('(6): ' + minutesSecondsMilliSeconds() + ' ' + 
                    ID of removed attachment: ' + asyncResult.value);
            }
        });


```

Même si  **addFileAttachmentAsync** démarre avant **removeAttachmentAsync**, comme  **addFileAttachmentAsync** est asynchrone, les appels `write` et **removeAttachmentAsync** peuvent démarrer avant que **addFileAttachmentAsync** ne se termine. Le cas échéant, `attachmentID` reste **undefined** et vous obtenez une erreur pour l’appel **removeAttachmentAsync**, comme dans la sortie suivante :




```
 (4): 46:18:245 attachmentID is: undefined
Error executing code: Sys.ArgumentException: Sys.ArgumentException: Value does not fall within the expected range. Parameter name: attachmentId
 (2): 46:18:255 ID of added attachment: 0
 (3): 46:18:262 Finishing addFileAttachmentAsync callback method.
```

Une façon d’éviter cela est de vérifier que  `attachmentID` est défini avant d’appeler **removeAttachmentAsync**. Une autre façon est d’initier l’appel  **removeAttachmentAsync** à partir de la méthode de rappel de **addFileAttachmentAsync**, comme le montre l’exemple suivant :




```js
var attachmentURI = "https://webserver/picture.png";
var attachmentID;

function minutesSecondsMilliSeconds()
{
    var d = new Date();
    return d.getMinutes() + ":" + d.getSeconds() + ":" + d.getMilliseconds();
}

Office.context.mailbox.item.addFileAttachmentAsync(
        attachmentURI,
        'Welcome document',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write('(1) ' + minutesSecondsMilliSeconds() + ' ' + 
                    asyncResult.error.message);
            }
            else {
                attachmentID = asyncResult.value;
                write('(2) ' + minutesSecondsMilliSeconds() + ' ' + 
                    'ID of added attachment: ' + attachmentID);

                // Move the write and removeAttachmentAsync calls here 
                // inside the addFileAttachmentAsync callback, after the 
                // attaching has succeeded.
                write ('(4): ' + minutesSecondsMilliSeconds() + ' ' + 
                    'attachmentID is: ' + attachmentID);

                Office.context.mailbox.item.removeAttachmentAsync(
                    attachmentID,
                    { asyncContext: null },
                    function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed){
                            write('(5) ' + minutesSecondsMilliSeconds() + ' ' + 
                                asyncResult.error.message);
                        }
                        else {
                            write('(6) ' + minutesSecondsMilliSeconds() + ' ' + 
                                'ID of removed attachment: ' + attachmentID);
                        }
                    });
            }

            write('(3) ' + minutesSecondsMilliSeconds() + ' ' + 
                'Finishing addFileAttachmentAsync callback method.');
        });

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Voici un exemple de sortie :




```
(2) 49:25:775 ID of added attachment: 1
(4) 49:25:782 attachmentID is: 1
(3) 49:25:783 Finishing addFileAttachmentAsync callback method.
(6) 49:25:789 ID of removed attachment: 1
```

Remarque : le rappel pour  **removeAttachmentAsync** est imbriqué à l’intérieur du rappel pour **addFileAttachmentAsync**. Comme  **addFileAttachmentAsync** et **removeAttachmentAsync** sont asynchrones, la dernière ligne dans le rappel pour **addFileAttachmentAsync** peut être exécutée avant la fin du rappel pour **removeAttachmentAsync**.


## Ressources supplémentaires



- [Créer des compléments Outlook pour les formulaires de composition](../outlook/compose-scenario.md)
    
- [Programmation asynchrone dans des compléments Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    


