
# Méthode TableBinding.addRowsAsync
Ajoute des lignes et des valeurs à un tableau.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|TableBindings|
|**Dernière modification dans **|1.1|

```js
bindingObj.addRowsAsync(rows, [,options], callback);
```


## Paramètres

_rows_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type : **Array**

&nbsp;&nbsp;&nbsp;&nbsp;Tableau de tableaux qui contient une ou plusieurs lignes de données à ajouter au tableau. Obligatoire.
    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type : **object**

&nbsp;&nbsp;&nbsp;&nbsp;Spécifie les [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants.
    
&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Type : **array, boolean, null, number, object, string ou non défini**<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet [AsyncResult](../../reference/shared/asyncresult.md) sans être modifié. Facultatif.<br/><br/>

_callback_<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Type :  **object**
    
&nbsp;&nbsp;&nbsp;&nbsp;Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type [AsyncResult](../../reference/shared/asyncresult.md). Facultatif.



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _rows_|**array**|Tableau de tableaux qui contient une ou plusieurs lignes de données à ajouter au tableau. Obligatoire.||
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants.||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **addRowsAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Renvoie toujours **undefined** car il n’existe aucun objet ni aucune donnée à récupérer.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Remarques

Le succès ou l’échec d’une opération **addRowsAsync** est atomique. En d’autres termes, l’ensemble de l’opération d’ajout de lignes doit réussir ; sinon, l’opération est complètement restaurée (en outre, la propriété **AsyncResult.status** qui est renvoyée au rappel signale un échec) :


- Chaque ligne du tableau que vous transmettez en tant qu’argument _data_ doit avoir le même nombre de colonnes que le tableau mis à jour. Sinon, toute l’opération échoue.
    
- Chaque ligne et chaque cellule du tableau doit ajouter correctement cette ligne et cette cellule au tableau dans la ou les nouvelles lignes ajoutées. S’il est impossible de définir une ligne ou une cellule pour une raison quelconque, toute l’opération échoue.
    
 **Remarques supplémentaires pour Excel Online**

Le nombre total de cellules dans la valeur transmise au paramètre _rows_ ne peut pas dépasser 20 000 dans un appel unique à cette méthode.


## Exemple




```js
function addRowsToTable() {
    Office.context.document.bindings.getByIdAsync("myBinding", function (asyncResult) {
        var binding = asyncResult.value;
        binding.addRowsAsync([["6", "k"], ["7", "j"]]);
    });
}

```




## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
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
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad|
|1.1|Prise en charge supplémentaire de l’écriture de données de tableau dans les compléments pour Access.|
|1.0|Introduit|
