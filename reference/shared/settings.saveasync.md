
# Méthode Settings.saveAsync
Fait persister la copie en mémoire du conteneur de propriétés des paramètres dans le document.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, PowerPoint, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Paramètres|
|**Dernière modification dans **|1.1|

```js
Office.context.document.settings.saveAsync(callback);
```


## Paramètres



_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type :  **object**

&nbsp;&nbsp;&nbsp;&nbsp;Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**. Facultatif.

    



## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **saveAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Renvoie toujours **undefined** car il n’existe aucun objet ni aucune donnée à récupérer.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Remarques

Tous les paramètres précédemment enregistrés par un complément sont chargés lorsqu’il est initialisé. Ainsi, pendant la durée de la session, il vous suffit d’employer les méthodes [set](../../reference/shared/settings.set.md) et [get](../../reference/shared/settings.get.md) pour utiliser la copie en mémoire du conteneur des propriétés des paramètres. Pour conserver les paramètres et pour qu’ils soient disponibles lors de la prochaine utilisation du complément, utilisez la méthode **saveAsync**.


 >**Remarque** : la méthode **saveAsync** fait persister le conteneur des propriétés des paramètres en mémoire dans le fichier de document. Cependant, les modifications apportées au fichier de document sont enregistrées uniquement lorsque l’utilisateur (ou le paramètre **AutoRecover**) enregistre le document dans le système de fichiers.

La méthode [refreshAsync](../../reference/shared/settings.refreshasync.md) est utile uniquement dans les scénarios de co-création (qui sont seulement pris en charge dans Word), lorsque d’autres instances du même complément peuvent modifier les paramètres, et ces modifications doivent être rendues disponibles sur toutes les instances.


## Exemple




```js
function persistSettings() {
    Office.context.document.settings.saveAsync(function (asyncResult) {
        write('Settings saved with status: ' + asyncResult.status);
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
