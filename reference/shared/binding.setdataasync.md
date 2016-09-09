
# Méthode Binding.setDataAsync
Écrit des données dans la section liée du document représenté par l’objet de liaison spécifié.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Disponible dans les [ensembles de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**Dernière modification dans TableBindings**|1.1|

```js
bindingObj.setDataAsync(data [, options] ,callback);
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _data_|<table><tr><td><b>string</b></td><td>Excel, Excel Online, Word et Word Online uniquement</td></tr><tr><td><b>tableau</b> (tableau de tableaux : matrice, « matrix »)</td><td>Excel et Word uniquement</td></tr><tr><td>
  <a href="https://msdn.microsoft.com/en-us/library/office/fp161002">
  <b>TableData</b></a></td><td>Access, Excel et Word uniquement</td></tr><tr><td><b>HTML</b></td><td>Word et Word Online uniquement</td></tr><tr><td><b>Office Open XML</b></td><td>Word uniquement</td></tr></table>|Données à définir dans la sélection actuelle. Requis.|**Modifié dans :** 1.1. La prise en charge des composants de contenu pour Access exige l’ensemble de ressources requis **TableBinding** 1.1 ou ultérieur.|
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants||
| _coercionType_|**[CoercionType](../../reference/shared/coerciontype-enumeration.md)**|Indique comment forcer le type des données définies. ||
| _colonnes_|**tableau de chaînes**| Spécifie les noms de colonne.|**Ajouté dans :** v1.1.Uniquement pour les liaisons de tableaux dans les compléments de contenu pour Access.|
| _Objet Rows_|**Office.TableRange.ThisRow**|Spécifie la chaîne prédéfinie « thisRow » pour définir les données dans la ligne actuellement sélectionnée. |**Ajouté dans :** v1.1.Uniquement pour les liaisons de tableaux dans les compléments de contenu pour Access.|
| _startColumn_|**number**|Spécifie la colonne de départ de base zéro pour un sous-ensemble des données. |Uniquement pour les liaisons de tableau ou de matrice. S’il est omis, les données sont définies à partir de la première colonne.|
| _startRow_|**number**|Spécifie la ligne de départ de base zéro pour un sous-ensemble des données dans la liaison. |Uniquement pour les liaisons de tableau ou de matrice. S’il est omis, les données sont définies à partir de la première ligne.|
| _tableOptions_|**object**|Pour le tableau inséré, liste de paires clé-valeur qui spécifient les [options de mise en forme de tableau](../../docs/excel/format-tables-in-add-ins-for-excel.md), comme la ligne d’en-tête, le nombre total de lignes et les lignes à bandes. |**Ajouté dans :** v1.1. **Pris en charge dans :** Excel.|
| _cellFormat_|**object**|Pour le tableau inséré, liste de paires clé-valeur qui spécifient la plage de cellules, lignes ou colonnes et la [mise en forme de cellule](../../docs/excel/format-tables-in-add-ins-for-excel.md) à appliquer à cette plage.|**Ajouté dans** v1.1. **Pris en charge dans :** Excel, Excel Online.|
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **setDataAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Renvoie toujours **undefined** car il n’existe aucun objet ni aucune donnée à récupérer.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Remarques

La valeur transmise pour _data_ contient les données à écrire dans la liaison. Le type de valeur transmis détermine ce qui sera écrit, comme le décrit le tableau suivant.



|**Valeur _data_**|**Données écrites**|
|:-----|:-----|
|Valeur **string**|Du texte brut ou tout élément dont le type peut être forcé en type **string** est écrit.|
|Tableau de tableaux (« matrice »)|Les données sous forme de tableau sans en-têtes seront écrites. Par exemple, pour écrire des données sur trois lignes dans deux colonnes, vous pouvez transmettre un tableau comme suit : ` [["R1C1", "R1C2"], ["R2C1", "R2C2"], ["R3C1", "R3C2"]]`Pour écrire une seule colonne de trois lignes, transmettez un tableau comme suit : `[["R1C1"], ["R2C1"], ["R3C1"]]`|
|Objet [TableData](../../reference/shared/tabledata.md)|Un tableau avec des en-têtes est écrit.|
En outre, ces actions (spécifiques aux applications) s’appliquent lors de l’écriture de données dans une liaison.

 **Pour Word**, le paramètre _data_ spécifié est écrit sur la liaison comme suit :



|**Valeur _data_**|**Données écrites**|
|:-----|:-----|
|Valeur **string**|Le texte spécifié est écrit.|
|Tableau de tableaux (« matrice ») ou objet **TableData**|Un tableau Word est écrit.|
|HTML|Le code HTML spécifié est écrit.
 >**Important**  Si le code HTML que vous écrivez n’est pas valide, Word ne déclenche aucune erreur. Word écrit autant de code HTML que possible et omet les données non valides.

|
|Office Open XML (« Open XML »)|Le code XML spécifié est écrit.|  **Pour Excel**, le paramètre _data_ spécifié est écrit sur la liaison comme suit :



|**Valeur _data_**|**Données écrites**|
|:-----|:-----|
|Valeur **string**|Le texte spécifié est inséré en tant que valeur de la première cellule liée. Vous pouvez également spécifier une formule valide pour l’ajouter à la cellule liée. Par exemple, la définition du paramètre _data_ sur `"=SUM(A1:A5)"` totalisera les valeurs de la plage spécifiée. Toutefois, après avoir défini une formule sur la cellule liée, vous ne pouvez pas lire la formule ajoutée (ni les formules préexistantes) à partir de la cellule liée. Si vous appelez la méthode [Binding.getDataAsync](../../reference/shared/binding.getdataasync.md) sur la cellule liée pour en lire les données, la méthode peut renvoyer uniquement les données affichées dans la cellule (le résultat de la formule).|
|Tableau de tableaux ("matrix") et la forme correspond exactement à la forme de la liaison spécifiée|L’ensemble de lignes et colonnes est écrit. Vous pouvez également spécifier un tableau de tableaux contenant des formules valides pour les ajouter aux cellules liées. Par exemple, la définition du paramètre _data_ sur `[["=SUM(A1:A5)","=AVERAGE(A1:A5)"]]` ajoutera ces deux formules à une liaison contenant deux cellules. Comme lors de la définition d’une formule sur une cellule liée unique, vous ne pouvez pas lire les formules ajoutées (ni les formules préexistantes) à partir de la liaison avec la méthode **Binding.getDataAsync** ; celle-ci renvoie uniquement les données affichées dans les cellules liées.|
|Objet **TableData** et la forme du tableau correspond à celle du tableau lié.|L’ensemble spécifié de lignes et/ou d’en-têtes est écrit, si aucune autre donnée dans les cellules environnantes ne sera écrasée. **Remarque :** si vous spécifiez des formules dans l’objet **TableData** que vous transmettez au paramètre _data_, vous risquez d’obtenir des résultats différents de ceux que vous attendez, en raison de la fonctionnalité d’Excel « Colonnes calculées », qui duplique automatiquement les formules dans une colonne. Pour contourner ce problème lorsque vous souhaitez écrire un paramètre _data_ contenant des formules vers une table liée, spécifiez les données sous forme de tableau de tableaux (au lieu de les spécifier sous forme d’objet **TableData**) et définissez le paramètre _coercionType_ sur **Microsoft.Office.Matrix** ou « matrix ».|
 **Remarques supplémentaires pour Excel Online**


- Le nombre total de cellules dans la valeur transmise au paramètre _data_ ne peut pas dépasser 20 000 dans un appel unique à cette méthode.
    
- Le nombre de _groupes de mise en forme_ transmis au paramètre _cellFormat_ ne peut pas dépasser 100. Un groupe de mise en forme se compose d’un ensemble de mises en forme appliquées à une plage de cellules donnée. Par exemple, l’appel suivant transmet deux groupes de mise en forme au paramètre _cellFormat_.
    
```js
  Office.select("bindings#myBinding").setDataAsync([['Berlin'],['Munich'],['Duisburg']],
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});

```

Dans tous les autres cas, une erreur est retournée.

La méthode **setDataAsync** écrit des données dans un sous-ensemble d’une liaison de tableau ou de matrice, si les paramètres facultatifs _startRow_ et _startColumn_ sont spécifiés, et s’ils indiquent une plage valide.


## Exemple




```js
function setBindingData() {
    Office.select("bindings#MyBinding").setDataAsync('Hello World!', function (asyncResult) { });
}
```

En spécifiant le paramètre facultatif _coercionType_, vous pouvez indiquer le type de données que vous souhaitez écrire dans une liaison. Par exemple, dans Word, si vous voulez écrire du contenu HTML dans une liaison de texte, vous pouvez spécifier le paramètre _coercionType_ en tant que `"html"` comme indiqué dans l’exemple suivant. Ce dernier utilise les balises HTML `<b>` pour mettre la chaîne « Hello » en gras.




```js
function writeHtmlData() {
    Office.select("bindings#myBinding").setDataAsync("<b>Hello</b> World!", {coercionType: "html"}, function (asyncResult) {
        if (asyncResult.status == "failed") {
            write('Error: ' + asyncResult.error.message);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Dans cet exemple, l’appel de **setDataAsync** transmet le paramètre _data_ en tant que tableau de tableaux (pour créer une seule colonne de trois lignes) et spécifie la structure de données avec `"matrix"` pour le paramètre _coercionType_.




```js
function writeBoundDataMatrix() {
    Office.select("bindings#myBinding").setDataAsync([['Berlin'],['Munich'],['Duisburg']],{ coercionType: "matrix" }, function (asyncResult) {
        if (asyncResult.status == "failed") {
            write('Error: ' + asyncResult.error.message);
        } else {
            write('Bound data: ' + asyncResult.value);
        }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Dans la fonction `writeBoundDataTable` de cet exemple, l’appel à **setDataAsync** transmet un objet _TableData_ pour le paramètre **data** (pour écrire trois colonnes et trois lignes) et spécifie la structure de données avec `"table"` pour le paramètre _coercionType_. 

Dans la fonction `updateTableData`, l’appel à **setDataAsync** transmet à nouveau un objet _TableData_ pour le paramètre **data**, mais avec une seule colonne, un nouvel en-tête et trois lignes, pour mettre à jour les valeurs de la dernière colonne du tableau créé à l’aide de la fonction `writeBoundDataTable`. Le paramètre facultatif de base zéro _startColumn_ est spécifié avec 2 pour remplacer les valeurs de la troisième colonne du tableau.




```js
function writeBoundDataTable() {
    // Create a TableData object.
    var myTable = new Office.TableData();
    myTable.headers = ['First Name', 'Last Name', 'Grade'];
    myTable.rows = [['Kim', 'Abercrombie', 'A'], ['Junmin','Hao', 'C'],['Toni','Poe','B']];

    // Set myTable in the binding.
    Office.select("bindings#myBinding").setDataAsync(myTable, { coercionType: "table" }, 
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                write('Error: '+ asyncResult.error.message);
        } else {
            write('Bound data: ' + asyncResult.value);
        }
    });
}

// Replace last column with different data.
function updateTableData() {
     var newTable = new Office.TableData();
     newTable.headers = ["Gender"];
     newTable.rows = [["M"],["M"],["F"]];
     Office.select("bindings#myBinding").setDataAsync(newTable, { coercionType: "table", startColumn:2 }, 
         function (asyncResult) {
             if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                 write('Error: '+ asyncResult.error.message);
         } else {
            write('Bound data: ' + asyncResult.value);
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
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|MatrixBindings, TableBindings, TextBindings|
|**Niveau d’autorisation minimal**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.1|<ul><li>Dans des compléments pour Access, l’écriture de données de tableau est désormais prise en charge.</li><li>Dans les compléments pour Excel, la <a href="http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33(Office.15).aspx" target="_blank">définition de la mise en forme lorsque vous écrivez des données dans une liaison de tableau</a> est désormais prise en charge à l’aide des paramètres facultatifs <span class="parameter" sdata="paramReference">tableOptions</span> et <span class="parameter" sdata="paramReference">cellFormat</span>.</li></ul>|
|1,0|Introduit|

## Voir aussi



#### Autres ressources


[Lier des régions dans un document ou une feuille de calcul](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)
