
# Méthode TableBinding.setFormatsAsync
Définit ou met à jour la mise en forme des éléments et données spécifiés dans le tableau lié.

|||
|:-----|:-----|
|**Hôtes :**|Excel|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Pas dans un ensemble|
|**Ajouté dans**|1.1|

```
bindingObj.setFormatsAsync(cellFormat [,options] , callback);
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _cellFormat_|**tableau**|Tableau contenant des objets JavaScript indiquant les cellules à cibler et la mise en forme à leur appliquer. Obligatoire.||
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel passée à la méthode **goToByIdAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Renvoie toujours **undefined** car il n’existe aucun objet ni aucune donnée à récupérer lors de la définition des formats.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Remarques

 **Spécification du paramètre cellFormat**

Utilisez le paramètre _cellFormat_ pour définir ou modifier les valeurs de mise en forme de cellule, telles que la largeur, la hauteur, la police, l’arrière-plan, l’alignement, etc. La valeur que vous transmettez pour le paramètre _cellFormat_ est de type **array** et contient la liste des objets JavaScript qui spécifient les cellules à cibler (`cells:`) et les mises en forme (`format:`) à appliquer à ces cellules.

Chaque objet JavaScript présent dans le tableau _cellFormat_ se présente comme suit :

 `{cells:{`_cell_range_`}, format:{`_format_definition_`}}`

La propriété `cells:` indique la plage que vous souhaitez mettre en forme à l’aide de l’une des valeurs suivantes :


**Plages prises en charge dans la propriété cells**


|**Paramètres de la plage de cellules**|**Description**|
|:-----|:-----|
| `{row: i}`|Spécifie la plage qui s’étend jusqu’à la ligne de données i dans le tableau.|
| `{column: i}`|Spécifie la plage qui s’étend jusqu’à la colonne de données i dans le tableau.|
| `{row: i, column: j}`|Spécifie la plage de cellules à partir de la ligne i jusqu’à la colonne de données j dans le tableau.|
| `Office.Table.All`|Spécifie le tableau entier, y compris les totaux, les données et les en-têtes de colonne (le cas échéant).|
| `Office.Table.Data`|Spécifie uniquement les données du tableau (sans les en-têtes ni les totaux).|
| `Office.Table.Headers`|Spécifie uniquement la ligne d’en-tête.|


La propriété `format:` spécifie les valeurs qui correspondent à un ensemble de paramètres disponibles dans la boîte de dialogue **Format de cellule** dans Excel (cliquez avec le bouton droit de la souris et sélectionnez **Format de cellule** ou cliquez sur **Accueil**  >  **Format**  >  **Format de cellule**).

Vous devez spécifier la valeur de la propriété `format:` sous la forme d’une liste de paires _nom de propriété_ - _valeur_ dans un littéral d’objet JavaScript. Le _nom de propriété_ indique le nom de la propriété de mise en forme à définir, tandis que la _valeur_ spécifie la valeur de la propriété. Vous pouvez spécifier plusieurs valeurs pour un format donné, comme la couleur et la taille de la police. Voici trois exemples de valeurs de la propriété `format:` :




```
//Set cells: font color to green and size to 15 points.
format: {fontColor : "green", fontSize : 15}
```




```
//Set cells: border to dotted blue.
format: {borderStyle: "dotted", borderColor: "blue"}
```




```
//Set cells: background to red and alignment to centered.
format: {backgroundColor: "red", alignHorizontal: "center"}
```

Vous pouvez indiquer des formats numériques en spécifiant la chaîne « code » de format numérique dans la propriété `numberFormat:`. Les chaînes de format numérique que vous pouvez spécifier correspondent à celles que vous pouvez définir dans Excel à l’aide de la catégorie **Personnalisée** sous l’onglet **Nombre** de la boîte de dialogue **Format de cellule**. L’exemple suivant montre comment mettre en forme un nombre en tant que pourcentage à deux décimales :




```
format: {numberFormat:"0.00%"}
```

Pour plus de détails, voir comment [créer un format numérique personnalisé](http://office.microsoft.com/en-us/excel-help/create-or-delete-a-custom-number-format-HA102749035.aspx?CTT=1#BM1).



 **Spécification d’une cible unique**

L’exemple suivant montre une valeur _cellFormat_ qui définit la couleur de la police de la ligne d’en-tête en rouge.




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: Office.Table.Headers, format: {fontColor: "red"}}], 
    function (asyncResult){});
```

 **Spécification de plusieurs cibles**

La méthode **setFormatsAsync** peut prendre en charge la mise en forme de plusieurs cibles dans le tableau lié dans un même appel de fonction. Pour ce faire, vous devez transmettre une liste d’objets dans le tableau _cellFormat_ pour chaque cible que vous souhaitez mettre en forme. Par exemple, la ligne de code suivante permet de définir la couleur jaune pour la police de la première ligne, ainsi qu’une bordure blanche et du texte gras dans la quatrième cellule de la troisième ligne.




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});
```

Pour définir la mise en forme des tables lors de l’écriture des données, utilisez les paramètres facultatifs _tableOptions_ et _cellFormat_ des méthodes [Document.setSelectedDataAsync](http://msdn.microsoft.com/library/4c1e13e9-b61a-47df-836c-3ca9aba4ca1c%28Office.15%29.aspx) et [TableBinding.setDataAsync](http://msdn.microsoft.com/library/5b6ecf6f-c57f-4c0d-9605-59daee8fde13%28Office.15%29.aspx).

La définition de la mise en forme à l’aide des paramètres facultatifs des méthodes **Document.setSelectedDataAsync** et **TableBinding.setDataAsync** ne fonctionne que lorsque vous définissez pour la première fois la mise en forme lors de l’écriture des données. Pour apporter des modifications de mise en forme après l’écriture de données, appliquez les méthodes suivantes :


- Pour mettre à jour la mise en forme des cellules, comme la couleur et le style de police, utilisez la méthode **TableBinding.setFormatsAsync** (cette méthode).
    
- Pour mettre à jour les options de table, comme les lignes à bandes et les boutons de filtre, appliquez la méthode [TableBinding.setTableOptions](../../reference/shared/binding.tablebinding.settableoptionsasync.md).
    
- Pour effacer la mise en forme, appliquez la méthode [TableBinding.clearFormats](../../reference/shared/binding.tablebinding.clearformatsasync.md).
    
 **Remarques supplémentaires pour Excel Online**

Le nombre de _groupes de mise en forme_ transmis au paramètre _cellFormat_ ne peut pas dépasser 100. Un groupe de mise en forme se compose d’un ensemble de mises en forme appliquées à une plage de cellules donnée. Par exemple, l’appel suivant transmet deux groupes de mise en forme au paramètre _cellFormat_.




```js
Office.select("bindings#myBinding).setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});

```

Pour plus d’informations et d’exemples, voir la rubrique relative à la [mise en forme des tables dans les compléments pour Excel](../../docs/excel/format-tables-in-add-ins-for-excel.md).


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**||**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|v||v|v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|Pas dans un ensemble.|
|**Niveau d’autorisation minimal**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel dans Office pour iPad.|
|1.1|Introduit|
