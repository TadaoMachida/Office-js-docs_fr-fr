
# Événement Office.initialize
Se produit quand l’environnement d’exécution est chargé et que le complément est prêt à interagir avec l’application et le document hébergé. 

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Dernière modification dans **|1.1|

```js
Office.initialize = function (reason) {/* initialization code */}
```


## Remarques

Le paramètre _reason_ de la fonction de détecteur d’événements **initialize** renvoie une valeur d’énumération [InitializationReason](../../reference/shared/initializationreason-enumeration.md) qui indique comment l’initialisation s’est produite. Un complément du volet Office ou de contenu peut être initialisé de deux façons :


- L’utilisateur vient de l’insérer à partir de la section **Compléments utilisés récemment** de la liste déroulante **Complément** sous l’onglet **Insertion** du ruban dans l’application hôte Office, ou à partir de la boîte de dialogue **Insérer un complément**.
    
- L’utilisateur a ouvert un document qui contient déjà le complément.
    

 >**Remarque** : le paramètre reason de la fonction de détecteur d’événements **initialize** renvoie uniquement une valeur d’énumération **InitializationReason** pour les compléments du volet Office et de contenu. Il ne renvoie aucune valeur pour les compléments Outlook.


## Exemple

Vous pouvez utiliser la valeur de **InitializationEnumeration** pour implémenter une autre logique quand le complément est inséré pour la première fois par opposition au moment où il fait déjà partie du document. L’exemple suivant illustre une logique simple qui utilise la valeur du paramètre _reason_ pour indiquer la façon dont le complément du volet Office ou de contenu a été initialisé.


```js
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, code specific to the add-in can run.
    // Display initialization reason.
    if (reason == "inserted")
    write("The add-in was just inserted.");

    if (reason == "documentOpened")
    write("The add-in is already part of the document.");
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet événement est pris en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cet événement.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|**OWA pour périphériques**|**Outlook pour Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||v||||
|**Excel**|v|v|v|||
|**Outlook**|v|v||v|v|
|**PowerPoint**|v|v|v|||
|**Projet**|v|||||
|**Word**|v|v|v|||

|||
|:-----|:-----|
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|De contenu Outlook, du volet Office|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint Online.|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire pour l’initialisation de compléments de contenu pour Access.|
|1.0|Introduit|
