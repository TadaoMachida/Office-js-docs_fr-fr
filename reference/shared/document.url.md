
# Propriété Document.url
Obtient l’URL du document actuellement ouvert dans l’application hôte.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Project, Word|
|**Dernière modification dans **|1.1|

```
var docUrl = Office.context.document.url;
```


## Valeur renvoyée

URL du document. Renvoie **null** si l’URL n’est pas disponible.


## Remarques

 **Important :** la propriété **url** renvoie les informations qui peuvent contenir des informations d’identification personnelle (PII) dans le nom du document et l’emplacement où il est stocké. Si vous devez stocker ou transmettre ces informations, veillez à le faire dans un format chiffré.


## Exemple




```
function displayDocumentUrl() {
    write(Office.context.document.url);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette propriété est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette propriété.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Projet**|v|||
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge





****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de Word Online.|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire des compléments de contenu pour Access.|
|1.0|Introduit|
