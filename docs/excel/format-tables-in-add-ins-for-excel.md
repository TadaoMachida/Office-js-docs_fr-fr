
# Mettre en forme les tableaux dans les compléments pour Excel


Cet article répertorie les différentes fonctionnalités de l’API de mise en forme et explique comment les utiliser. Dans cette version, vous pouvez spécifier par programme la mise en forme des cellules et d’autres options dans les tableaux uniquement (pas pour les structures de données  **Office.CoercionType.Text** ou **Office.CoercionType.Matrix**) et dans les compléments uniquement pour Excel. Pour définir une mise en forme avec votre complément, procédez comme suit :

- L’utilisateur sélectionne le tableau (ou l’emplacement où insérer un tableau par programme), puis votre complément peut appeler la méthode  **Document.setSelectedDataAsync** sur ce tableau pour définir la mise en forme.

- Ou, si le classeur contient déjà des tableaux liés (ou si votre complément utilise l’une des méthodes « addFrom » de l’objet [Bindings](../../reference/shared/bindings.bindings.md) pour créer des tableaux liés quand il est initialisé), votre complément peut appeler la méthode **Binding.setDataAsync** sur ces tableaux liés pour définir la mise en forme.
    
>**Important :** pour utiliser ces méthodes nouvelles et mises à jour afin de mettre en forme des tableaux dans les compléments pour Excel, votre projet de complément doit [utiliser ou être mis à jour pour utiliser Office.js version 1.1 ou ultérieure](../../docs/develop/update-your-javascript-api-for-office-and-manifest-schema-version.md).

## Spécification de la mise en forme

Pour spécifier la mise en forme que vous souhaitez définir, créez un littéral d’objet JavaScript qui contient une ou plusieurs paires clé-valeur. Vous pouvez combiner une série de paramètres de mise en forme dans une liste dans l’objet JavaScript. Par exemple : 


```js
var myFormat = {fontStyle:"bold", width:"autoFit", borderColor:"purple"};
```

Pour appliquer la mise en forme, passez l’objet JavaScript à l’une des méthodes qui prend en charge la mise en forme de données et d’autres fonctionnalités du tableau.

Vous pouvez utiliser la mise en forme de deux manières :


- la première fois que votre complément écrit des données dans une sélection ou une liaison, en spécifiant les paramètres facultatifs _cellFormat_ ou _tableOptions_ dans l’objet _options_ transmis aux méthodes [Document.setSelectedDataAysnc](../../reference/shared/document.setselecteddataasync.md) ou [Binding.setDataAsync](../../reference/shared/binding.setdataasync.md) ;
    
- après la configuration initiale de la mise en forme, vous pouvez [effacer ou mettre à jour la mise en forme](#effacer-ou-mettre-à-jour-la-mise-en-forme) à l’aide de l’une des nouvelles méthodes prévues à cet effet.
    

## Utilisation de paramètres facultatifs avec des méthodes de définition de données

Pour les tableaux liés, vous pouvez utiliser les paramètres facultatifs suivants pour spécifier la mise en forme lors de la définition de données à l’aide des méthodes **Document.setSelectedData** ou **Binding.setDataAsync** : _tableOptions_ et _cellFormat_


### Paramètre facultatif tableOptions

Utilisez le paramètre facultatif  _tableOptions_ pour spécifier des styles de tableau par défaut, activer et désactiver certaines fonctionnalités de tableau comme **Ligne d’en-tête**,  **Ligne des totaux** et **Lignes à bandes**. La valeur que vous passez en tant que paramètre  _tableOptions_ est un objet JavaScript qui contient une liste de paires clé-valeur. Par exemple,


```js
tableOptions: {bandedRows: true, filterButton: false, style:"TableStyleMedium3"};
```


### Le paramètre facultatif cellFormat

Utilisez le paramètre facultatif  _cellFormat_ pour modifier les valeurs de mise en forme de cellule, telles que la largeur, la hauteur, la police, l’arrière-plan, l’alignement, etc. La valeur que vous passez en tant que paramètre _cellFormat_ est un tableau qui contient la liste des objets JavaScript qui spécifient les cellules à cibler et les formats qui s’y appliquent. Par exemple :


```js
cellFormat: 
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: Office.Table.Headers, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}]
```

Vous pouvez combiner plusieurs paires `cells:` et `format:` dans le tableau _cellFormat_ afin de réduire le nombre d’appels de fonction requis pour appliquer la mise en forme.


#### cells

Utilisez `cells:` pour spécifier la plage de colonnes, de lignes et de cellules à laquelle appliquer la mise en forme.


**Plages prises en charge dans les valeurs de cellules**


|**paramètres de plage de cellules**|**Description**|
|:-----|:-----|
| `{row: i}`|Spécifie la plage qui s’étend jusqu’à la ligne de données i dans le tableau.|
| `{column: i}`|Spécifie la plage qui s’étend jusqu’à la colonne de données i dans le tableau.|
| `{row: i, column: j}`|Spécifie la plage de cellules à partir de la ligne i jusqu’à la colonne de données j dans le tableau.|
| `Office.Table.All`|Spécifie le tableau entier, y compris les totaux, les données et les en-têtes de colonne (le cas échéant).|
| `Office.Table.Data`|Spécifie uniquement les données du tableau (sans les en-têtes ni les totaux).|
| `Office.Table.Headers`|Spécifie uniquement la ligne d’en-tête.|

#### format

Utilisez le paramètre `format:` pour spécifier la mise en forme à appliquer à la plage définie avec le paramètre `cells:` en tant que liste de paires clé-valeur JavaScript. Pour obtenir la liste des valeurs prises en charge, voir [Clés et valeurs de mise en forme prises en charge](#clés-et-valeurs-de-mise-en-forme-prises-en-charge).

 **Limites de spécification de la mise en forme pour Excel Online**

Lors de la définition de la mise en forme dans Excel Online, le nombre de _groupes de mise en forme_ passés au paramètre _cellFormat_ ne peut pas dépasser 100. Un seul groupe de mise en forme se compose d’un ensemble de mises en forme appliqué à une plage de cellules donnée. (En d’autres termes, tout ce qui est spécifié dans l’un des littéraux d’objet `cells:` dans le tableau est passé à_cellFormat_.) Par exemple, l’appel suivant passe deux groupes de mise en forme à _cellFormat_.




```js
Office.context.document.setSelectedDataAsync(
    {cellFormat:[{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}]}, 
    function (asyncResult){});
```


#### Application de paramètres facultatifs

Dans cette version, seules les méthodes **Document.setSelectedDataAsync** et **TableBinding.setDataAsync** méthodes prennent en charge l’écriture de données et la définition de mise en forme pour les tableaux dans le même appel à l’aide des paramètres facultatifs _tableOptions_ et _cellFormat_. Dans les exemples suivants, la valeur `tableData` passée au premier paramètre de chaque méthode (le paramètre _data_) doit être un objet [TableData](../../reference/shared/tabledata.md) qui contient la définition du tableau et des données à écrire.

 **Exemple Document.setSelectedDataAsync**




```js
Office.context.document.setSelectedDataAsync(tableData, 
    {tableOptions: {headerRow:false}, 
        cellFormat: [{cells: Office.Table.Headers, format: {fontColor: "yellow"}}]}, 
    function (asyncResult) {});
```

 **Exemple TableBinding.setDataAsync**




```js
Office.select("bindings#myBinding").setDataAsync(tableData, 
    {tableOptions: {headerRow:false}, 
        cellFormat: [{cells: Office.Table.Headers, format: {fontColor: "yellow"}}]}, 
    function (asyncResult) {});
```

 >**Remarque** : l’appel à `Office.select("bindings#myBinding")` suppose qu’une liaison nommée `myBinding` existe déjà dans la feuille de calcul.


## Mise à jour et suppression de la mise en forme


Lorsque vous définissez la mise en forme avec les paramètres facultatifs _cellFormat_ et _tableOptions_ des méthodes **Document.setSelectedDataAsync** ou **TableBinding.setDataAsync**, ils définiront la mise en forme uniquement la première fois que vous les appelez. Pour mettre à jour ou désactiver la mise en forme, vous devez utiliser trois nouvelles méthodes de l’objet **TableBinding** : **setFormatsAsync**, **setTableOptionsAsync** et **clearFormatsAsync**.


### Mise à jour de la mise en forme

La méthode [TableBinding.setFormatsAsync](../../reference/shared/binding.tablebinding.setformatsasync.md) est utilisée uniquement pour mettre à jour la mise en forme des cellules, telle que la largeur, la hauteur, la police, l’arrière-plan et l’alignement. Elle admet _cellFormat_ comme paramètre obligatoire :


```js
Office.select("bindings#myBinding").setFormatsAsync(
    [{cells: {row: 1}, format: {fontColor: "yellow"}}, 
        {cells: {row: 3, column: 4}, format: {borderColor: "white", fontStyle: "bold"}}], 
    function (asyncResult){});
```

La méthode [TableBinding.setTableOptionsAsync](../../reference/shared/binding.tablebinding.settableoptionsasync.md) est utilisée uniquement pour mettre à jour les options de tableau, telles que les lignes à bandes et les boutons de filtre. Elle admet _tableOptions_ comme paramètre obligatoire :




```js
var tableOptions = {bandedRows: true, filterButton: false, style: "TableStyleMedium3"}; 

Office.select("bindings#myBinding").setTableOptionsAsync(tableOptions, function(asyncResult){});
```


### Suppression de la mise en forme

La méthode [TableBinding.clearFormatsAsync](../../reference/shared/binding.tablebinding.clearformatsasync.md) permet de désactiver toute la mise en forme dans le tableau. Elle admet le paramètre facultatif _asyncContext_ et une fonction de rappel facultative :


```js
Office.select("bindings#myBinding").clearFormatsAsync();
```


## Clés et valeurs de mise en forme prises en charge


Les tableaux suivants répertorient les paires clé-valeur prises en charge que vous pouvez passer dans les paramètres _cellFormat_ ou _tableOptions_.

Pour les valeurs de `format:`, les paramètres que vous pouvez spécifier sont un sous-ensemble des paramètres disponibles dans la boîte de dialogue **Format de cellule** (cliquez avec le bouton droit de la souris et sélectionnez **Format de cellule** ou **Format** > **Format de cellule** sous l’onglet **Accueil** du ruban). Pour les valeurs `tableOptions:`, les paramètres sont ceux disponibles dans les groupes **Options de style de tableau** et **Styles de tableau** sous l’onglet **Outils de tableau** |**Création** du ruban.


 >**Important** :  Les méthodes de l’API de mise en forme prennent en charge uniquement les options et les valeurs résumées ci-dessous. Si vous spécifiez d’autres options ou valeurs de mise en forme, la gestion des erreurs est non définie. La gestion non définie des erreurs n’est pas nécessairement cohérente sur toutes les plateformes prises en charge ; vous ne devez pas développer de compléments basés sur l’un des effets secondaires de la gestion non définie des erreurs pour toute plateforme spécifique. Toutefois, la gestion non définie des erreurs ne doit nuire ni à l’état ni à l’interface utilisateur de votre complément, ni aux documents avec lesquels il interagit.


**Alignement**


|**Touche**|**Valeurs**|**Notes**|
|:-----|:-----|:-----|
| `alignHorizontal:`|"general" \| "left" \| "center" \| "right" \| "fill" \| "justify" \| "center across selection" \| "distributed"|En cas d’association avec une valeur de retrait, seules les combinaisons suivantes sont prises en charge :<br/><br/><ul><li><code>alignHorizontal: "left"</code> et <code>indentLeft: \<value\></code></li></ul><ul><li><code>alignHorizontal: "right"</code> et <code>indentRight: \<value\></code></li></ul><ul><li><code>alignHorizontal: "distributed"</code> et <code>indentDistributed: \<value\></code></li></ul>|
| `alignVertical:`|"top" \| "center" \| "bottom" \| "justify" \| "distributed"||



**Arrière-plan**


|**Touche**|**Valeurs**|**Notes**|
|:-----|:-----|:-----|
| `backgroundColor:`|"none" \| \<Tous les noms de couleurs prédéfinis\> \| #RRGGBB|Noms de couleurs prédéfinis :<br/><br/>"black", "blue", "gray", "green", "orange", "pink", "purple", "red", "teal", "turquoise", "violet", "white", "yellow"|



**Bordure**


|**Touche**|**Valeurs**|**Notes**|
|:-----|:-----|:-----|
| `borderStyle:`|"none" \| \<Tous les noms de style de bordure prédéfinis\>|Noms de style de bordure prédéfinis :<br/><br/>"dash dot", "dash dot dot", "dashed", "dotted", "double", "hair", "medium dash dot", "medium dash dot dot", "medium dashed", "medium", "slant dash dot", "thick", "thin"<br/><br/><Tous les noms de styles de bordure prédéfinis> (Revient à spécifier des styles de bordure à l’aide de la présélection **Contour** et **Intérieur** sous l’onglet **Bordure** de la boîte de dialogue **Format de cellule**.)<br/><br/> **Remarque :** Excel 2013 prend en charge le rendu des 13 styles de bordure prédéfinis. Toutefois, Excel Online ne prend pas en charge tous les styles de bordure. Le tableau suivant décrit le rendu utilisé pour chaque style bordure lorsque vous ouvrez la feuille de calcul dans Excel Online.<br/><br/><table><tr><th>Excel 2013</th><th>Excel Online</th></tr><tr><td>"dash dot"</td><td>"dash dot"</td></tr><tr><td>dashed (1 pixel)</td><td>"dash dot dot"</td></tr><tr><td>dotted (1 pixel)</td><td>"dashed"</td></tr><tr><td>dotted (1 pixel)</td><td>"dotted"</td></tr><tr><td>dashed (1 pixel)</td><td>"double"</td></tr><tr><td>double (3 pixels)</td><td>"hair"</td></tr><tr><td>solid (1 pixel)</td><td>"medium dash dot"</td></tr><tr><td>dashed (2 pixels)</td><td>"medium dash dot dot"</td></tr><tr><td>dotted (2 pixels)</td><td>"medium dashed"</td></tr><tr><td>dashed (2 pixels)</td><td>"medium"</td></tr><tr><td>solid (2 pixels)</td><td>"slant dash dot"</td></tr><tr><td>dashed (2 pixels)</td><td>"thick"</td></tr><tr><td>solid (3 pixels)</td><td>solid (1 pixel)</td></tr></table>|
| `borderColor:`|"automatic" \| \<Tous les noms de couleurs prédéfinis\> \| #RRGGBB|<Tous les noms de styles de bordure prédéfinis>|
| `borderTopStyle:`|"none" \| \<Tous les noms de style de bordure prédéfinis\>|<Tous les noms de styles de bordure prédéfinis>|
| `borderTopColor:`|"automatic" \| \<Tous les noms de couleurs prédéfinis\> \| #RRGGBB|<Tous les noms de styles de bordure prédéfinis>|
| `borderBottomStyle:`|"none" \| \<Tous les noms de style de bordure prédéfinis\>|<Tous les noms de styles de bordure prédéfinis>|
| `borderBottomColor:`|"automatic" \| \<Tous les noms de couleurs prédéfinis\> \| #RRGGBB|<Tous les noms de styles de bordure prédéfinis>|
| `borderLeftStyle:`|"none" \| \<Tous les noms de style de bordure prédéfinis\>|<Tous les noms de styles de bordure prédéfinis>|
| `borderLeftColor:`|"automatic" \| \<Tous les noms de couleurs prédéfinis\> \| #RRGGBB|<Tous les noms de styles de bordure prédéfinis>|
| `borderRightStyle:`|"none" \| \<Tous les noms de style de bordure prédéfinis\>|<Tous les noms de styles de bordure prédéfinis>|
| `borderRightColor:`|"automatic" \| \<Tous les noms de couleurs prédéfinis\> \| #RRGGBB|<Tous les noms de styles de bordure prédéfinis>|
| `borderOutlineStyle:`|"none" \| \<Tous les noms de style de bordure prédéfinis\>|<Tous les noms de styles de bordure prédéfinis>|
| `borderOutlineColor:`|"automatic" \| \<Tous les noms de couleurs prédéfinis\> \| #RRGGBB|<Tous les noms de styles de bordure prédéfinis>|
| `borderInlineStyle:`|"none" \| \<Tous les noms de style de bordure prédéfinis\>|S’applique uniquement aux bordures intérieures dans la plage spécifiée. (Revient à spécifier des styles de bordure à l’aide de la présélection **Intérieur** sous l’onglet **Bordure** de la boîte de dialogue **Format de cellule**.)|
| `borderInlineColor:`|"automatic" \| \<Tous les noms de couleurs prédéfinis\> \| #RRGGBB|#RRGGBB |



**Largeur, hauteur de la cellule et renvoi à la ligne**


|**Touche**|**Valeurs**|
|:-----|:-----|
| `width:`|"auto fit" \|  **Nombre**|
| `height:`|"auto fit" \|  **Nombre**|
| `wrapping:`|**Boolean**|



**Police**


|**Touche**|**Valeurs**|**Notes**|
|:-----|:-----|:-----|
| `fontFamily:`|\<Tous les noms de polices disponibles\>|<Tous les noms de polices disponibles>|
| `fontStyle:`|"regular" \| "italic" \| "bold" \| "bold"|**Remarque** : Au moment de la publication, le paramètre `fontStyle:` défini sur "italic", puis sur "bold" (ou vice versa) par la suite, se comporte comme une association de ces deux paramètres. Autrement dit, si, par exemple, vous définissez d’abord "italic" et "bold" ensuite, le résultat sera "bold italic". Pour définir l’italique ou gras _uniquement_ sur une plage précédemment définie en gras ou en italique, vous devez d’abord définir `fontStyle:"regular"` pour effacer la mise en forme précédente.|
| `fontSize:`|**Nombre**||
| `fontUnderlineStyle:`|"none" \| "single" \| "double" \| "single accounting" \| "single accounting"||
| `fontColor:`|"automatic" \| \<Tous les noms de couleurs prédéfinis\> \| #RRGGBB||
| `fontDirection:`|"context" \| "left-to-right" \| "left-to-right"|Excel Online ne prend actuellement pas en charge l’affichage du texte de droite à gauche. Toutefois, si votre complément définit `fontDirection:` sur "right-to-left" lorsqu’il est exécuté dans Excel Online, ce paramètre de mise en forme est enregistré dans le fichier du classeur et s’affiche correctement lorsque le classeur est ouvert dans Excel pour ordinateur de bureau.|
| `fontStrikethrough:`|**Boolean**||
| `fontSuperscript:`|**Booléen**||
| `fontSubScript:`|**Booléen**||
| `fontNormal:`|**Boolean**|Définit la police, le style de police, la taille et les effets sur le style normal. Cela réinitialise la mise en forme des caractères de la cellule avec les valeurs par défaut. Équivaut à cocher la case **Police normale** sur l’onglet **Police** de la boîte de dialogue **Format de cellules**.|



**Retrait**


|**Touche**|**Valeurs**|**Notes**|
|:-----|:-----|:-----|
| `indentLeft:`|**Nombre**|Lorsque cette clé est combinée à une valeur de retrait, seule la combinaison suivante est prise en charge :<br/><br/><ul><li><code>alignHorizontal: "left"</code> et <code>indentLeft: \<value\></code></li></ul>|
| `indentRight:`|**Nombre**|Lorsque cette clé est combinée à une valeur de retrait, seule la combinaison suivante est prise en charge :<br/><br/><ul><li><code>alignHorizontal: "right"</code> et <code>indentRight: \<value\></code></li></ul>|
| `indentDistributed:`|**Nombre**|Lorsque cette clé est combinée à une valeur de retrait, seule la combinaison suivante est prise en charge :<br/><br/><ul><li><code>alignHorizontal: "distributed"</code> et <code>indentDistributed: \<value\></code></li></ul>|



**Format des nombres**


|**Touche**|**Valeurs**|**Notes**|
|:-----|:-----|:-----|
| `numberFormat:`|**String**|Pour spécifier le format des nombres, utilisez une chaîne de format de nombre personnalisée. Par exemple, pour spécifier deux décimales avec une virgule comme séparateur de milliers, vous devez spécifier :<br/><br/> `numberFormat:"#,###.00"`<br/><br/>Ce sont les mêmes chaînes de format personnalisé que vous pouvez [créer avec la catégorie de format personnalisé sous l’onglet Nombre dans la boîte de dialogue Format de cellules](http://office.microsoft.com/en-us/excel-help/create-or-delete-a-custom-number-format-HA102749035.aspx?CTT=1).<br/><br/> **Conseil :** Vous pouvez voir à quoi ressemble une chaîne de format pour une catégorie standard dans la boîte de dialogue **Format de cellules** dans Excel en suivant les étapes suivantes :<br/><br/><ol><li>Sélectionnez une catégorie de format standard, par exemple <span class="ui">Devise</span>, dans la liste <b>Catégorie</b>.</li><li>Définissez les options de format dans la partie droite de la boîte de dialogue.</li><li>Sélectionnez la catégorie <b>Personnalisation</b> pour afficher la chaîne de format en haut de la liste <b>Type</b>.</li></ol>|



**Options de tableau**


|**Touche**|**Valeurs**|**Notes**|
|:-----|:-----|:-----|
| `style:`|"none" \| \<Tous les noms de style de tableaux prédéfinis\>|Tous les noms de style de tableaux prédéfinis :<br/><br/>"TableStyleLight1", "TableStyleLight2", "TableStyleLight3", "TableStyleLight4", "TableStyleLight5", "TableStyleLight6", "TableStyleLight7", "TableStyleLight8", "TableStyleLight9", "TableStyleLight10", "TableStyleLight11", "TableStyleLight12", "TableStyleLight13", "TableStyleLight14", "TableStyleLight15", "TableStyleLight16", "TableStyleLight17", "TableStyleLight18", "TableStyleLight19", "TableStyleLight20", "TableStyleLight21", "TableStyleMedium1", "TableStyleMedium2", "TableStyleMedium3", "TableStyleMedium4", "TableStyleMedium5", "TableStyleMedium6", "TableStyleMedium7", "TableStyleMedium8", "TableStyleMedium9", "TableStyleMedium10", "TableStyleMedium11", "TableStyleMedium12", "TableStyleMedium13", "TableStyleMedium14", "TableStyleMedium15", "TableStyleMedium16", "TableStyleMedium17", "TableStyleMedium18", "TableStyleMedium19", "TableStyleMedium20", "TableStyleMedium21", "TableStyleMedium22", "TableStyleMedium23", "TableStyleMedium24", "TableStyleMedium25", "TableStyleMedium26", "TableStyleMedium27", "TableStyleMedium28", "TableStyleDark1", "TableStyleDark2", "TableStyleDark3", "TableStyleDark4", "TableStyleDark5", "TableStyleDark6", "TableStyleDark7", "TableStyleDark8", "TableStyleDark9", "TableStyleDark10", "TableStyleDark11"<br/><br/>Pour voir à quoi un style de tableau ressemble, insérez un tableau dans Excel, sur l’onglet **Outils de tableau** \| **Création**, choisissez la liste déroulante  **Styles rapides**, puis sélectionnez un style prédéfini. L’info-bulle relative au style correspond à l’une des valeurs figurant dans la liste ci-dessus.|
| `headerRow:`|**Boolean**||
| `firstColumn:`|**Booléen**||
| `filterButton:`|**Booléen**||
| `totalRow:`|**Booléen**||
| `lastColumn:`|**Booléen**||
| `bandedRows:`|**Booléen**||
| `bandedColumns:`|**Boolean**||
