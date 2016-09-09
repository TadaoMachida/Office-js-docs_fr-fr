
# Méthode Binding.getDataAsync
Retourne les données contenues dans la liaison.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Disponible dans les [ensembles de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**Dernière modification dans TableBindings**|1.1|

```
bindingObj.getDataAsync([, options] , callback );
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants||
| _coercionType_|**[CoercionType](../../reference/shared/coerciontype-enumeration.md)**|Indique comment forcer le type des données définies. ||
| _valueFormat_|[ValueFormat](../../reference/shared/valueformat-enumeration.md)|Spécifie si les valeurs, telles que les nombres et les dates, sont renvoyées avec leur mise en forme appliquée.||
| _filterType_|[FilterType](../../reference/shared/filtertype-enumeration.md)|Spécifie s’il faut appliquer un filtre quand les données sont récupérées.||
| _Objet Rows_|**Office.TableRange.ThisRow**| Spécifie la chaîne prédéfinie « thisRow » pour obtenir des données dans la ligne actuellement sélectionnée.|Uniquement pour les liaisons de tableau dans les compléments de contenu pour Access.|
| _startRow_|**number**|Pour les liaisons de tableau ou de matrice, spécifie la ligne de départ de base zéro pour un sous-ensemble des données de la liaison. ||
| _startColumn_|**number**|Pour les liaisons de tableau ou de matrice, spécifie la colonne de départ de base zéro pour un sous-ensemble des données de la liaison. ||
| _rowCount_|**number**|Pour les liaisons de tableau ou de matrice, spécifie le nombre de lignes décalées par rapport à _startRow_. ||
| _columnCount_|**number**|Pour les liaisons de tableau ou de matrice, spécifie le nombre de colonnes décalées par rapport à _startColumn_.||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **Binding.getDataAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Accédez aux valeurs dans la liaison spécifiée. Si le paramètre _coercionType_ est spécifié (et si l’appel a réussi), les données sont renvoyées au format décrit dans la rubrique relative à l’énumération [CoercionType](../../reference/shared/coerciontype-enumeration.md).|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Remarques

Si un paramètre facultatif est omis, la valeur par défaut suivante est utilisée (quand elle s’applique au type et au format des données).



|**Paramètre**|**Par défaut**|
|:-----|:-----|
| _coercionType_|Type d’origine, non forcé, de la liaison.|
| _valueFormat_|Données non mises en forme.|
| _filterType_|Toutes les valeurs (non filtrées).|
| _startRow_|Première ligne.|
| _startColumn_|Première colonne.|
| _rowCount_|Toutes les lignes.|
| _columnCount_|Toutes les colonnes.|
Quand elle est appelée à partir de [MatrixBinding](../../reference/shared/binding.matrixbinding.md) ou [TableBinding](../../reference/shared/binding.tablebinding.md), la méthode **getDataAsync** renvoie un sous-ensemble des valeurs liées si les paramètres facultatifs _startRow_, _startColumn_, _rowCount_ et _columnCount_ sont spécifiés (et s’ils spécifient une plage à la fois contigüe et valide).


## Exemple




```
function showBindingData() {
    Office.select("bindings#MyBinding").getDataAsync(function (asyncResult) {
        write(asyncResult.value)
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```



Il existe une différence majeure de comportement entre les valeurs `"table"` et `"matrix"`_coercionType_ avec la méthode **Binding.getDataAsync** pour les données formatées avec des lignes d’en-tête, comme illustré dans les deux exemples suivants. Ces exemples de code indiquent les fonctions de gestionnaire d’événements pour l’événement [Binding.SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md).

Si vous spécifiez `"table"` pour _coercionType_, la propriété [TableData.rows](../../reference/shared/tabledata.rows.md) (`result.value.rows` dans l’exemple de code suivant) renvoie un tableau qui contient uniquement les lignes du corps du tableau. Ainsi, sa ligne 0 sera la première ligne qui n’est pas une ligne d’en-tête dans le tableau.




```js
function selectionChanged(evtArgs) { 
    Office.select("bindings#TableTranslate").getDataAsync({ coercionType: 'table', startRow: evtArgs.startRow, startCol: 0, rowCount: 1, columnCount: 1 },  
        function (result) { 
            if (result.status == 'succeeded') { 
                write("Image to find: " + result.value.rows[0][0]); 
            } 
            else 
                write(result.error.message); 
    }); 
}     
// Function that writes to a div with id='message' on the page. 
function write(message){ 
    document.getElementById('message').innerText += message; 
}
```

Toutefois, si vous spécifiez `"matrix"` pour _coercionType_, `result.value` dans l’exemple de code suivant renvoie un tableau qui contient l’en-tête du tableau en ligne 0. Si l’en-tête du tableau contient plusieurs lignes, celles-ci sont toutes incluses dans la matrice `result.value` en tant que lignes séparées avant les lignes du corps du tableau.




```js
function selectionChanged(evtArgs) { 
    Office.select("bindings#TableTranslate").getDataAsync({ coercionType: 'matrix', startRow: evtArgs.startRow, startCol: 0, rowCount: 1, columnCount: 1 },  
        function (result) { 
            if (result.status == 'succeeded') { 
                write("Image to find: " + result.value[1][0]); 
            } 
            else 
                write(result.error.message); 
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


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
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
|1.1|Prise en charge supplémentaire des liaisons de tableau dans les compléments pour Access.|
|1.0|Introduit|

## Voir aussi



#### Autres ressources


[Lier des régions dans un document ou une feuille de calcul](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)
