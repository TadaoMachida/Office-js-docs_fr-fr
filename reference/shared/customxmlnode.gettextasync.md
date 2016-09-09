
# Méthode CustomXmlNode.getTextAsync
Obtient de manière asynchrone le texte d’un nœud XML dans une partie XML personnalisée.

|||
|:-----|:-----|
|**Hôtes :**|Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Ajouté dans**|1.2|

```js
customXmlNodeObj.getTextAsync([asyncContext,]callback(asyncResult);
```


## Paramètres



|**Nom**|**Type**|**Description**|
|:-----|:-----|:-----|
| _asyncContext_|**object**|Facultatif. Un objet défini par l’utilisateur disponible sur la propriété asyncCesult de l’objet [AsyncResult](../../reference/shared/asyncresult.md). Utilisez ce paramètre pour indiquer un objet ou une valeur à **AsyncResult** lorsque le rappel est une fonction nommée.|
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.|

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **getTextAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Accéder à une **chaîne** qui contient le texte interne des nœuds référencés.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Indique si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_. Cette propriété renvoie undefined si _asyncContext_ n’a pas été défini.|

## Exemple

Découvrez comment obtenir la valeur de texte d’un nœud dans une partie XML personnalisée.


```js
// Get the built-in core properties XML part by using its ID. This results in a call to Word.
Office.context.document.customXmlParts.getByIdAsync("{6C3C8BC8-F283-45AE-878A-BAB7291924A1}", function (getByIdAsyncResult) {
    
    // Access the XML part.
    var xmlPart = getByIdAsyncResult.value;
    
    // Add namespaces to the namespace manager. These two calls result in two calls to Word.
    xmlPart.namespaceManager.addNamespaceAsync('cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', function () {
        xmlPart.namespaceManager.addNamespaceAsync('dc', 'http://purl.org/dc/elements/1.1/', function () {

            // Get XML nodes by using an Xpath expression. This results in a call to Word.
            xmlPart.getNodesAsync("/cp:coreProperties/dc:title", function (getNodesAsyncResult) {
                
                // Get the first node returned by using the Xpath expression. 
                var node = getNodesAsyncResult.value[0];
                
                // Get the text value of the node and use the asyncContext. This results in a call to Word. 
                // The results are logged to the browser console.
                node.getTextAsync({asyncContext: "StateNormal"}, function (getTextAsyncResult) {
                   console.log("Text of the title element = " + getTextAsyncResult.value;
                   console.log("The asyncContext value = " + getTextAsyncResult.asyncContext;
                });
            });
        });
    });
});
```


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|CustomXmlParts|
|**Niveau d’autorisation minimal**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Ajout de getTextAsync.|
