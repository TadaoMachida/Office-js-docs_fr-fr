# Objet TableSort (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Excel pour iOS, Office 2016_

Gère les opérations de tri des objets Table.

## Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|matchCase|bool|Indique si la casse a influé sur le dernier tri du tableau. En lecture seule.|
|méthode|string|Dernière méthode de classement des caractères chinois utilisée pour trier le tableau. En lecture seule. Les valeurs possibles sont les suivantes : PinYin, StrokeCount|

## Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|champs|[SortField](sortfield.md)|Dernières conditions utilisées pour trier le tableau. En lecture seule.|

## Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[apply(fields: SortField[], matchCase: bool, method: string)](#applyfields-sortfield-matchcase-bool-method-string)|void|Effectue une opération de tri.|
|[clear()](#clear)|void|Efface le tri actuellement appliqué au tableau. Même si le classement du tableau n’est pas modifié, l’état des boutons d’en-tête est rétabli.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|
|[reapply()](#reapply)|void|Applique à nouveau les paramètres actuels de tri au tableau.|

## Détails des méthodes


### apply(fields: SortField[], matchCase: bool, method: string)
Effectue une opération de tri.

#### Syntaxe
```js
tableSortObject.apply(fields, matchCase, method);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|champs|SortField[]|Liste des conditions de tri.|
|matchCase|bool|Facultatif. Indique si la casse influe sur le classement des chaînes.|
|méthode|string|Facultatif. Méthode de classement utilisée pour les caractères chinois.  Les valeurs possibles sont les suivantes : PinYin, StrokeCount|

#### Retourne
void

#### Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.apply([ 
            {
                key: 2,
                ascending: true
            },
        ], true);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### clear()
Efface le tri actuellement appliqué au tableau. Même si le classement du tableau n’est pas modifié, l’état des boutons d’en-tête est rétabli.

#### Syntaxe
```js
tableSortObject.clear();
```

#### Paramètres
Aucun

#### Retourne
void

#### Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.clear();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});

### load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### Syntax
```js
object.load(param);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### Renvoie
void

### reapply()
Applique à nouveau les paramètres actuels de tri au tableau.

#### Syntaxe
```js
tableSortObject.reapply();
```

#### Paramètres
Aucun

#### Retourne
void

####Exemples
```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    table.sort.reapply();   
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});