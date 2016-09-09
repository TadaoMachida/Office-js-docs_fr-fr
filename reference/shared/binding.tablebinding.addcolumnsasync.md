
# Méthode TableBinding.addColumnsAsync
Ajoute des colonnes et des valeurs à un tableau.

|||
|:-----|:-----|
|**Hôtes :**|Excel, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Dernière modification dans **|1,0|

```
bindingObj.addColumnsAsync(data [, options], callback);
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _data_|**array** ou [TableData](../../reference/shared/tabledata.md)|Tableau de tableaux (matrice, « matrix ») ou objet **TableData** contenant une ou plusieurs lignes de données à ajouter au tableau. Requis.||
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants.||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **addColumnsAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Renvoie toujours **undefined** car il n’existe aucun objet ni aucune donnée à récupérer.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Remarques

Pour ajouter une ou plusieurs colonnes en spécifiant les valeurs des données et des en-têtes, transmettez un objet **TableData** en tant que paramètre _data_. Pour ajouter une ou plusieurs colonnes en spécifiant uniquement les données, transmettez un tableau de tableaux (matrice, « matrix ») pour le paramètre _data_.

Le succès ou l’échec d’une opération **addColumnAsync** est atomique. En d’autres termes, l’ensemble de l’opération d’ajout de colonnes doit réussir ; sinon, l’opération est complètement restaurée (en outre, la propriété **AsyncResult.status** qui est renvoyée au rappel signale un échec) :


- Chaque ligne du tableau que vous transmettez en tant qu’argument _data_ doit avoir le même nombre de lignes que le tableau mis à jour. Sinon, toute l’opération échoue.
    
- Chaque ligne et chaque cellule du tableau doit ajouter correctement cette ligne ou cette cellule au tableau dans la ou les nouvelles colonnes ajoutées. S’il est impossible de définir une ligne ou une cellule pour une raison quelconque, toute l’opération échoue.
    
- Si vous transmettez un objet **TableData** en tant qu’argument de données, le nombre de lignes d’en-tête doit correspondre à celui du tableau en cours de mise à jour.
    
**Remarques supplémentaires pour Excel Online**

Le nombre total de cellules dans l’objet **TableData** transmis au paramètre _data_ ne peut pas dépasser 20 000 dans un appel unique à cette méthode.


## Exemple

L’exemple suivant ajoute une seule colonne de trois lignes à un tableau lié ayant l’[id](../../reference/shared/binding.id.md)`"myTable"` en transmettant un objet **TableData** en tant qu’argument _data_ de la méthode **addColumnsAsync**. Pour que l’opération soit une réussite, le tableau en cours de mise à jour doit avoir trois lignes.


```js
// Add a column to a binding of type table by passing a TableData object.
function addColumns() {
    var myTable = new Office.TableData();
    myTable.headers = [["Cities"]];
    myTable.rows = [["Berlin"], ["Roma"], ["Tokyo"]];

    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addColumnsAsync(myTable);
    });
}
```

L’exemple suivant ajoute une seule colonne de trois lignes à un tableau lié ayant l’[id](../../reference/shared/binding.id.md)`myTable` en transmettant un tableau de tableaux (matrice, « matrix ») en tant qu’argument _data_ de la méthode **addColumnsAsync**. Pour que l’opération soit une réussite, le tableau en cours de mise à jour doit avoir trois lignes.




```js
// Add a column to a binding of type table by passing an array of arrays.
function addColumns() {
    var myTable = [["Berlin"], ["Roma"], ["Tokyo"]];

    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addColumnsAsync(myTable);
    });
}
```


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|TableBindings|
|**Niveau d’autorisation minimal**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.0|Introduit|
