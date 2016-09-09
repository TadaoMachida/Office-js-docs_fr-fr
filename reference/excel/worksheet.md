# Objet Worksheet (interface API JavaScript pour Excel)

Une feuille de calcul Excel est une grille de cellules. Elle peut contenir des données, des tableaux, des graphiques, etc.

## Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|id|string|Renvoie une valeur qui permet d’identifier la feuille de calcul de façon unique dans un classeur donné. La valeur de l’identificateur reste identique, même lorsque la feuille de calcul est renommée ou déplacée. Les valeurs changent à chaque ouverture de session du fichier. En lecture seule.|
|name|string|Nom complet de la feuille de calcul.|
|position|int|Position de la feuille de calcul au sein du classeur (sur une base zéro).|
|visibility|string|Visibilité de la feuille de calcul. Les valeurs possibles sont les suivantes : Visible (visible), Hidden (masquée), VeryHidden (très masquée).|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|charts|[ChartCollection](chartcollection.md)|Renvoie une collection de graphiques qui font partie de la feuille de calcul. En lecture seule.|
|protection|[WorksheetProtection](worksheetprotection.md)|Renvoie un objet de protection de feuille pour une feuille de calcul. En lecture seule.|
|tables|[TableCollection](tablecollection.md)|Collection de tableaux qui font partie de la feuille de calcul. En lecture seule.|

## Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[activate()](#activate)|void|Active la feuille de calcul dans l’interface utilisateur Excel.|
|[delete()](#delete)|void|Supprime la feuille de calcul du classeur.|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Renvoie l’objet de plage qui contient une cellule donnée sur la base des numéros de ligne et de colonne. La cellule peut se trouver en dehors des limites de ses plages parent, pour peu qu’elle reste dans la grille de la feuille de calcul.|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|Renvoie l’objet de plage spécifié par son nom ou son adresse.|
|[getUsedRange(valuesOnly: bool)](#getusedrangevaluesonly-bool)|[Range](range.md)|La plage utilisée est la plus petite plage qui englobe toutes les cellules auxquelles une valeur ou un format est affecté. Si la feuille de calcul est vide, cette fonction renvoie la cellule supérieure gauche.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails des méthodes


### activate()
Active la feuille de calcul dans l’interface utilisateur Excel.

#### Syntaxe
```js
worksheetObject.activate();
```

#### Paramètres
Aucun

#### Retourne
void

#### Exemples

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.activate();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### delete()
Supprime la feuille de calcul du classeur.

#### Syntaxe
```js
worksheetObject.delete();
```

#### Paramètres
Aucun

#### Retourne
void

#### Exemples

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.delete();
    return ctx.sync(); 
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getCell(row: number, column: number)
Renvoie l’objet de plage qui contient une cellule donnée sur la base des numéros de ligne et de colonne. La cellule peut se trouver en dehors des limites de ses plages parent, pour peu qu’elle reste dans la grille de la feuille de calcul.

#### Syntaxe
```js
worksheetObject.getCell(row, column);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|row|number|Numéro de ligne de la cellule à récupérer. Avec indice zéro.|
|column|number|Numéro de colonne de la cellule à récupérer. Avec indice zéro.|

#### Retourne
[Range](range.md)

#### Exemples

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var cell = worksheet.getCell(0,0);
    cell.load('address');
    return ctx.sync().then(function() {
        console.log(cell.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getRange(address: string)
Renvoie l’objet de plage spécifié par son nom ou son adresse.

#### Syntaxe
```js
worksheetObject.getRange(address);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|address|string|Facultatif. Adresse ou nom de la plage. Si cette propriété n’est pas définie, la plage de la feuille de calcul toute entière est renvoyée.|

#### Retourne
[Range](range.md)

#### Exemples
Cet exemple utilise l’adresse de la plage pour obtenir l’objet de la plage.

```js
Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeAddress = "A1:F8";
    var worksheet = ctx.workbook.worksheets.getItem(sheetName);
    var range = worksheet.getRange(rangeAddress);
    range.load('cellCount');
    return ctx.sync().then(function() {
        console.log(range.cellCount);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Cet exemple utilise une plage nommée pour obtenir l’objet de la plage.

```js

Excel.run(function (ctx) { 
    var sheetName = "Sheet1";
    var rangeName = 'MyRange';
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeName);
    range.load('address');
    return ctx.sync().then(function() {
        console.log(range.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getUsedRange(valuesOnly: bool)
La plage utilisée est la plus petite plage qui englobe toutes les cellules auxquelles une valeur ou un format est affecté. Si la feuille de calcul est vide, cette fonction renvoie la cellule supérieure gauche.

#### Syntaxe
```js
worksheetObject.getUsedRange(valuesOnly);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|valuesOnly|bool|Facultatif. Prend uniquement en compte les cellules avec des valeurs sous forme de cellules utilisées (ignore la mise en forme).|

#### Retourne
[Range](range.md)

#### Exemples

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    var usedRange = worksheet.getUsedRange();
    usedRange.load('address');
    return ctx.sync().then(function() {
            console.log(usedRange.address);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### Syntaxe
```js
object.load(param);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### Renvoie
void
### Exemples d’accès aux propriétés

Obtenir les propriétés de la feuille de calcul à partir du nom de la feuille

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.load('position')
    return ctx.sync().then(function() {
            console.log(worksheet.position);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Définir la position de la feuille de calcul 

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sheet1';
    var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
    worksheet.position = 2;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

