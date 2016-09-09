
# Méthode File.getSliceAsync
Retourne la section spécifiée.

|||
|:-----|:-----|
|**Hôtes :**|PowerPoint, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Fichier|
|**Ajouté dans**|1,0|

```js
File.getSliceAsync(sliceIndex, callback);
```


## Paramètres


_sliceIndex_ <br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type : **number**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Spécifie l’index de base zéro de la section à récupérer. Obligatoire.<br/><br/>
    
_callback_ <br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type :  **object**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type [AsyncResult](../../reference/shared/asyncresult.md). Facultatif.
    

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **getSliceAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Accéder à l’objet [Slice](../../reference/shared/slice.md).|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible dans l’ensemble de ressources requis**|Fichier|
|**Niveau d’autorisation minimal**|[ReadDocument (ReadAllDocument requis pour accéder à Office OpenXML)](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint et Word dans Office pour iPad.|
|1.0|Introduit|
