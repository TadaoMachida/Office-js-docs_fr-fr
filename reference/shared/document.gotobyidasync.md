
# Méthode Document.goToByIdAsync
Accède à l’emplacement ou l’objet spécifié dans le document.

|||
|:-----|:-----|
|**Hôtes :**|Excel, PowerPoint, Word|
|**Disponible dans les ensembles de ressources requis**|Pas dans un ensemble|
|**Ajouté dans**|1.1|

```js
Office.context.document.goToByIdAsync(id, goToType, [,options], callback);
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _id_|**string** ou **number**|Identifiant de l’objet ou de l’emplacement à atteindre. Obligatoire.||
| _goToType_|[GoToType](../../reference/shared/gototype-enumeration.md)|Type d’emplacement à atteindre. Obligatoire.||
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants||
| _selectionMode_|[SelectionMode](../../reference/shared/selectionmode-enumeration.md)|Indique si l’emplacement spécifié par le paramètre _id_ est sélectionné (en surbrillance).|**Dans Excel :**<br/> **Office.SelectionMode.Selected** sélectionne tout le contenu de la liaison ou de l’élément nommé. <br/>**Office.SelectionMode.None** pour les liaisons de texte, sélectionne la cellule ; pour les liaisons de matrice, les liaisons de tableau et les éléments nommés, sélectionne la première cellule de données (pas la première cellule dans la ligne d’en-tête pour les tableaux).<br/><br/> **Dans PowerPoint :**<br/> **Office.SelectionMode.Selected** sélectionne le titre de la diapositive ou la première zone de texte sur la diapositive.<br/> **Office.SelectionMode.None** ne sélectionne rien.<br/><br/> **Dans Word :**<br/> **Office.SelectionMode.Selected** sélectionne tout le contenu de la liaison. <br/>**Office.SelectionMode.None** pour les liaisons de texte, déplace le curseur au début du texte ; pour les liaisons de matrice et de tableau, sélectionne la première cellule de données (pas la première cellule dans la ligne d’en-tête pour les tableaux).|
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel passée à la méthode **goToByIdAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Renvoyer l’affichage actuel.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Remarques

PowerPoint ne prend pas en charge la méthode **goToByIdAsync** dans les **Modes Masques**.


## Exemple

 **Accéder à une liaison en fonction de son ID (Word et Excel)**

Observez l’exemple suivant :


-  **Créer une liaison de tableau** en prenant la méthode [addFromSelectionAsync](../../reference/shared/bindings.addfromselectionasync.md) comme exemple.
    
-  **Indiquer cette liaison** comme étant celle à atteindre.
    
-  **Passer une fonction de rappel anonyme** qui renvoie le statut de l’opération au paramètre _callback_ de la méthode **goToByIdAsync**.
    
-  **Afficher la valeur** sur la page du complément.
    



```js
function gotoBinding() {
    //Create a new table binding for the selected table.
    Office.context.document.bindings.addFromSelectionAsync("table",{ id: "MyTableBinding" }, function (asyncResult) {
    if (asyncResult.status == "failed") {
              showMessage("Action failed with error: " + asyncResult.error.message);
           }
           else {
              showMessage("Added new binding with type: " + asyncResult.value.type +" and id: " + asyncResult.value.id);
           }
    });

    //Go to binding by id.
    Office.context.document.goToByIdAsync("MyTableBinding", Office.GoToType.Binding, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```



 **Accéder à un tableau dans une feuille de calcul (Excel)**

Observez l’exemple suivant :


-  **Indiquer le nom** du tableau à atteindre.
    
-  **Passer une fonction de rappel anonyme** qui renvoie le statut de l’opération au paramètre _callback_ de la méthode **goToByIdAsync**.
    
-  **Afficher la valeur** sur la page du complément.
    



```js
function goToTable() {
    Office.context.document.goToByIdAsync("Table1", Office.GoToType.NamedItem, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
    });
}
```



 **Accéder à la diapositive sélectionnée en fonction de son ID (PowerPoint)**

Observez l’exemple suivant :


-  **Obtenir l’ID** des diapositives sélectionnées à l’aide de la méthode [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md).
    
-  **Indiquer l’ID renvoyé** comme étant la diapositive à atteindre.
    
-  **Passer une fonction de rappel anonyme** qui renvoie le statut de l’opération au paramètre _callback_ de la méthode **goToByIdAsync**.
    
-  **Afficher la valeur** de l’objet JSON sous forme de chaîne renvoyé par `asyncResult.value`, qui contient des informations concernant les diapositives sélectionnées sur la page du complément.
    



```js
var firstSlideId = 0;
function gotoSelectedSlide() {
    //Get currently selected slide's id
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
    //Go to slide by id.
    Office.context.document.goToByIdAsync(firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```



 **Accéder à la diapositive en fonction de son index (PowerPoint)**

Observez l’exemple suivant :


-  **Indiquer l’index** de la diapositive à atteindre (première, dernière, précédente ou suivante).
    
-  **Passer une fonction de rappel anonyme** qui renvoie le statut de l’opération au paramètre _callback_ de la méthode **goToByIdAsync**.
    
-  **Afficher la valeur** sur la page du complément.
    



```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage("Navigation successful");
        }
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
|**PowerPoint**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|Pas dans un ensemble|
|**Niveau d’autorisation minimal**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint Online.|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Introduit|
