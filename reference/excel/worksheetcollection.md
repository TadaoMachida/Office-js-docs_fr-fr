# Objet WorksheetCollection (interface API JavaScript pour Excel)

Représente une collection d’objets de feuille de calcul qui font partie du classeur.

## Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|Items|[Worksheet[]](worksheet.md)|Collection d’objets de feuille de calcul. En lecture seule.|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
Aucun


## Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[add(name: string)](#addname-string)|[Feuille de calcul](worksheet.md)|Ajoute une nouvelle feuille de calcul au classeur. La feuille de calcul est ajoutée à la fin des feuilles de calcul existantes. Si vous souhaitez activer la feuille de calcul nouvellement ajoutée, appelez la méthode .activate() pour cette feuille.|
|[getActiveWorksheet()](#getactiveworksheet)|[Feuille de calcul](worksheet.md)|Obtient la feuille de calcul active du classeur.|
|[getItem(key: string)](#getitemkey-string)|[Feuille de calcul](worksheet.md)|Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails des méthodes


### add(name: string)
Ajoute une nouvelle feuille de calcul au classeur. La feuille de calcul est ajoutée à la fin des feuilles de calcul existantes. Si vous souhaitez activer la feuille de calcul nouvellement ajoutée, appelez la méthode .activate() pour cette feuille.

#### Syntaxe
```js
worksheetCollectionObject.add(name);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|name|string|Facultatif. Nom de la feuille de calcul à ajouter. Si cette propriété est définie, le nom doit être unique. Si cette propriété n’est pas définie, Excel détermine le nom de la nouvelle feuille de calcul.|

#### Retourne
[Feuille de calcul](worksheet.md)

#### Exemples

```js
Excel.run(function (ctx) { 
    var wSheetName = 'Sample Name';
    var worksheet = ctx.workbook.worksheets.add(wSheetName);
    worksheet.load('name');
    return ctx.sync().then(function() {
        console.log(worksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getActiveWorksheet()
Obtient la feuille de calcul active du classeur.

#### Syntaxe
```js
worksheetCollectionObject.getActiveWorksheet();
```

#### Paramètres
Aucun

#### Retourne
[Feuille de calcul](worksheet.md)

#### Exemples

```js
Excel.run(function (ctx) {  
    var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
    activeWorksheet.load('name');
    return ctx.sync().then(function() {
            console.log(activeWorksheet.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getItem(key: string)
Obtient un objet de feuille de calcul à l’aide de son nom ou de son ID.

#### Syntaxe
```js
worksheetCollectionObject.getItem(key);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|Key|string|Nom ou ID de la feuille de calcul.|

#### Retourne
[Feuille de calcul](worksheet.md)

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
```js
Excel.run(function (ctx) {
  var worksheets = ctx.workbook.worksheets;
  worksheets.load({"items" : "id, name"});
  return ctx.sync().then(function() {
    for (var i = 0; i < worksheets.items.length; i++)
    {
      console.log(worksheets.items[i].name);
      console.log(worksheets.items[i].id);
    }
  });
}).catch(function(error) {
  console.log("Error: " + error);
  if (error instanceof OfficeExtension.Error) {
    console.log("Debug info: " + JSON.stringify(error.debugInfo));
  }
});
```
