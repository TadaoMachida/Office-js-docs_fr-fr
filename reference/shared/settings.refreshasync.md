

# Méthode Settings.refreshAsync
Lit tous les paramètres persistants dans le document et actualise la copie de ces paramètres en mémoire pour le complément de contenu ou du volet Office.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, PowerPoint, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Paramètres|
|**Dernière modification dans **|1.1|

```js
Office.context.document.settings.refreshAsync(callback);
```


## Paramètres

_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type :  **object**

&nbsp;&nbsp;&nbsp;&nbsp;Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.

    



## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **refreshAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Accéder à un objet [Settings](../../reference/shared/settings.md) avec les valeurs actualisées.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Remarques

Cette méthode est utile dans les scénarios de co-création Word et PowerPoint lorsque plusieurs instances du même complément sont exécutées pour un même document. Comme chaque complément s’exécute en fonction d’une copie en mémoire des paramètres chargés à partir du document au moment où l’utilisateur l’a ouvert, les valeurs des paramètres utilisés par chaque utilisateur peuvent se désynchroniser. Cela peut se produire chaque fois qu’une instance du complément appelle la méthode [Settings.saveAsync](../../reference/shared/settings.saveasync.md) pour faire persister tous les paramètres de l’utilisateur concernant le document. L’appel de la méthode **refreshAsync** à partir du gestionnaire d’événements pour l’événement [settingsChanged](../../reference/shared/settings.settingschangedevent.md) du complément actualise les valeurs des paramètres pour tous les utilisateurs.

La méthode **refreshAsync** peut être appelée à partir des compléments créés pour Excel, mais comme elle ne prend pas en charge la co-création, il n’y a aucune raison de le faire.


## Exemple




```js
function refreshSettings() {
    Office.context.document.settings.refreshAsync(function (asyncResult) {
        write('Settings refreshed with status: ' + asyncResult.status);
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|Paramètres|
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint Online.|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire des paramètres personnalisés dans les compléments du contenu Access.|
|1.0|Introduit|
