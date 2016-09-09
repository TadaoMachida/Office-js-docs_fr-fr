
# Propriété Context.displayLanguage
Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Dernière modification dans **|1.1|

```
var myDisplayLanguage = Office.context.displayLanguage;
```


## Valeur renvoyée

Une **chaîne** au format de balise de langue du document RFC 1766, par exemple `en-US`.


## Remarques

La valeur **displayLanguage** reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier**  >  **Options**  >  **Langue** dans l’application hôte Office.

Dans les compléments de contenu pour les applications web Access, la propriété **displayLanguage** obtient la langue du complément (par exemple, « en-US »).


## Exemple




```js
function sayHelloWithDisplayLanguage() {
    var myDisplayLanguage = Office.context.displayLanguage;
    switch (myDisplayLanguage) {
        case 'en-US':
            write('Hello!');
            break;
        case 'en-NZ':
            write('G\'day mate!');
            break;
    }
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|**Outlook pour Mac**|
|:-----|:-----|:-----|:-----|:-----|
|**Access**||v|||
|**Excel**|v|v|v||
|**Outlook**|v|v||v|
|**PowerPoint**|v|v|v||
|**Projet**|v||||
|**Word**|v|v|v||

|||
|:-----|:-----|
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|De contenu, de volet de tâche, Outlook|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint Online.|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Ajout de l’accès à cette API dans les compléments de contenu pour Access.|
|1.0|Introduit|
