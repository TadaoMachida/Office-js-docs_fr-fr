
# Méthode Bindings.addFromSelectionAsync
Ajoute une liaison à la sélection actuelle dans le document.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**Dernière modification**|1.1|

```
bindingsObj.addFromSelectionAsync(bindingType [, options], callback);
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _bindingType_|[BindingType](../../reference/shared/bindingtype-enumeration.md)|Spécifie le type de l’objet de liaison à créer. Obligatoire. Renvoie **null** si le type spécifié ne peut pas être forcé sur l’objet sélectionné.||
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants.||
| _id_|**string**|Spécifie le nom unique à utiliser pour identifier le nouvel objet de liaison. Si aucun argument n’est transmis pour le paramètre _id_, le [Binding.id](../../reference/shared/binding.id.md) est généré automatiquement.||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **addFromSelectionAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Accéder à l’objet [Binding](../../reference/shared/binding.md) représentant la sélection spécifiée par l’utilisateur.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Remarques

Ajoute le type spécifié d’objet de liaison à la collection **Bindings** qui sera identifiée avec l’_id_ indiqué.


 >**Remarque**  Dans Excel, si vous appelez la méthode **addFromSelectionAsync** en transmettant le **Binding.id** d’une liaison existante, le [Binding.type](../../reference/shared/binding.type.md) de cette liaison est utilisé et son type ne peut pas être modifié en spécifiant une valeur différente pour le paramètre _bindingType_. Si vous devez utiliser un _id_ existant et modifier le _bindingType_, appelez d’abord la méthode [Bindings.releaseByIdAsync](../../reference/shared/bindings.releasebyidasync.md) pour libérer la liaison, puis appelez la méthode **addFromSelectionAsync** pour rétablir la liaison avec un nouveau type.


## Exemple

Ajoute une liaison [TextBinding](../../reference/shared/binding.textbinding.md) à la sélection active avec un identificateur **Binding.id** de « MyBinding ».


```js
function addBindingFromSelection() {
    Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'MyBinding' }, 
        function (asyncResult) {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    );
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|MatrixBindings, TableBindings, TextBindings|
|**Niveau d’autorisation minimal**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.1|Dans les compléments pour Excel, vous pouvez créer une liaison de tableau (en transmettant _Office.BindingType.Table_ pour **bindingType**) pour une plage de cellules qui contient des données tabulaires même lorsque les données n’ont pas été ajoutées à la feuille de calcul sous forme de tableau (à l’aide des commandes **Insérer**  >  **Tableaux**  > **Tableau** ou **Accueil**  >  **Styles**  >  **Mettre sous forme de tableau**).|
|1.1|Prise en charge supplémentaire de la liaison de tableau dans les compléments de contenu pour Access. |
|1.0|Introduit|
