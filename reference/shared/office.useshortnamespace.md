

# Méthode Office.useShortNamespace
Active ou désactive l’alias `Office` pour l’espace de noms `Microsoft.Office.WebExtension` complet.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Dernière modification dans **|1.1|

```js
Office.useShortNamespace(useShortcut);
```


## Paramètres



_useShortcut_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type :  **boolean**

    
valeur &nbsp;&nbsp;&nbsp;&nbsp;**true** pour utiliser le raccourci alias, ou valeur **false** pour le désactiver. La valeur par défaut est **true**.
    


## Exemple



```js
function startUsingShortNamespace() {
    if (typeof Office === 'undefined') {
        Microsoft.Office.WebExtension.useShortNamespace(true);
    }
    else {
        Office.useShortNamespace(true);
    }
    write('Office alias is now ' + typeof Office);
}

function stopUsingShortNamespace() {
    if (typeof Office === 'undefined') {
        Microsoft.Office.WebExtension.useShortNamespace(false);
    }
    else {
        Office.useShortNamespace(false);
    }
    write('Office alias is now ' + typeof Office);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

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
|1.1|Prise en charge supplémentaire de l’appel de cette méthode dans les compléments de contenu pour Access.|
|1.0|Introduit|
