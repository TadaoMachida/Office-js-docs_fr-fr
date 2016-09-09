
# Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook

L’interface de l’API JavaScript pour Office fournit des méthodes asynchrones ([getAsync](../../reference/outlook/Location.md) et [setAsync](../../reference/outlook/Location.md)) pour obtenir et définir l’emplacement d’un rendez-vous composé par l’utilisateur. Ces méthodes asynchrones sont disponibles uniquement pour les compléments de composition. Pour utiliser ces méthodes, assurez-vous que vous avez correctement configuré le manifeste du complément pour Outlook afin d’activer le complément dans des formulaires de composition, comme décrit dans la rubrique [Créer des compléments Outlook pour les formulaires de composition](../outlook/compose-scenario.md).

La propriété [location](../../reference/outlook/Office.context.mailbox.item.md) est disponible pour un accès en lecture dans les formulaires de lecture et de composition de rendez-vous. Dans un formulaire de lecture, vous pouvez accéder à la propriété directement à partir de l’objet parent, comme dans :




```js
item.location
```

Cependant, dans un formulaire de composition, comme l’utilisateur et votre complément peuvent insérer ou modifier l’emplacement en même temps, vous devez utiliser la méthode asynchrone  **getAsync** pour obtenir l’emplacement, comme indiqué ci-dessous :




```js
item.location.getAsync
```

La propriété  **location** est disponible pour l’accès en écriture uniquement dans les formulaires de composition de rendez-vous, mais pas dans les formulaires de lecture.

Comme avec la plupart des méthodes asynchrones dans l’interface API JavaScript pour Office,  **getAsync** et **setAsync** admettent des paramètres d’entrée facultatifs. Pour plus d’informations sur la spécification de ces paramètres d’entrée facultatifs, voir la rubrique [Programmation asynchrone dans des compléments Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md).


## Obtention de l’emplacement


Cette section présente un exemple de code qui obtient l’emplacement du rendez-vous que l’utilisateur compose, et affiche cet emplacement. Cet exemple de code suppose l’existence d’une règle dans le manifeste du complément qui active le complément dans un formulaire de composition pour un rendez-vous, comme indiqué ci-dessous.


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

Pour utiliser  **item.location.getAsync**, indiquez une méthode de rappel qui vérifie l’état et le résultat de l’appel asynchrone. Vous pouvez indiquer tous les arguments nécessaires à la méthode de rappel via le paramètre facultatif  _asyncContext_. Vous pouvez obtenir l’état, les résultats et toute erreur à l’aide du paramètre de sortie  _asyncResult_ du rappel. Si l’appel asynchrone aboutit, vous pouvez obtenir l’emplacement comme chaîne à l’aide de la propriété [AsyncResult.value](../../reference/outlook/simple-types.md).




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Définition de l’emplacement


Cette section présente un exemple de code qui définit l’emplacement du rendez-vous que l’utilisateur compose. Comme dans l’exemple précédent, cet exemple de code suppose l’existence d’une règle dans le manifeste de complément qui active le complément dans un formulaire de composition pour un rendez-vous.

Pour utiliser  **item.location.setAsync**, indiquez une chaîne de 255 caractères maximum dans le paramètre de données. Vous pouvez éventuellement fournir une méthode de rappel et tous les arguments pour la méthode de rappel dans le paramètre  _asyncContext_. Vous devez vérifier l’état, le résultat et tous les messages d’erreur dans le paramètre de sortie  _asyncResult_ du rappel. Si l’appel asynchrone aboutit, **setAsync** insère la chaîne d’emplacement spécifiée sous forme de texte brut, en écrasant tous les emplacements existants pour cet élément.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
                // Do whatever appropriate for your scenario
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Ressources supplémentaires



- [Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Obtention et définition de données d’élément Outlook dans des formulaires de lecture ou de composition](../outlook/item-data.md)
    
- [Créer des compléments Outlook pour les formulaires de composition](../outlook/compose-scenario.md)
    
- [Programmation asynchrone dans des compléments Office](../../docs/develop/asynchronous-programming-in-office-add-ins.md)
    
- [Obtenir, définir ou ajouter des destinataires lors de la composition d’un rendez-vous ou d’un message dans Outlook](../outlook/get-set-or-add-recipients.md)
    
- [Obtenir ou définir l’objet lors de la composition d’un rendez-vous ou d’un message dans Outlook](../outlook/get-or-set-the-subject.md)
    
- [Insérer des données dans le corps lors de la composition d’un rendez-vous ou d’un message dans Outlook](../outlook/insert-data-in-the-body.md)
    
- [Obtenir ou définir l’heure lors de la composition d’un rendez-vous dans Outlook](../outlook/get-or-set-the-time-of-an-appointment.md)
    
