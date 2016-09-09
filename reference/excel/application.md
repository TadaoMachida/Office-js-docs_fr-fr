# Objet Application (interface API JavaScript pour Excel)

Représente l’application Excel qui gère le classeur.

## Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|calculationMode|string|Renvoie le mode de calcul du classeur. En lecture seule. Les valeurs possibles sont les suivantes : `Automatic` Excel contrôle le recalcul, `AutomaticExceptTables` Excel contrôle le recalcul, mais ignore les modifications apportées aux tables, `Manual` le calcul est effectué lorsque l’utilisateur le demande.|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
Aucun


## Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|Recalcule tous les classeurs actuellement ouverts dans Excel.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails des méthodes


### calculate(calculationType: string)
Recalcule tous les classeurs actuellement ouverts dans Excel.

#### Syntaxe
```js
applicationObject.calculate(calculationType);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|calculationType|string|Spécifie le type de calcul à utiliser. Les valeurs possibles sont les suivantes : `Recalculate` (option par défaut), effectue le calcul normalement en appliquant toutes les formules du classeur, `Full` force le calcul intégral des données, `FullRebuild` force le calcul intégral des données et régénère les dépendances.|

#### Retourne
void

#### Exemples
```js
Excel.run(function (ctx) { 
    ctx.workbook.application.calculate('Full');
    return ctx.sync(); 
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
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, accepte un objet [loadOption](loadoption.md).|

#### Renvoie
void
### Exemples d’accès aux propriétés
```js
Excel.run(function (ctx) { 
    var application = ctx.workbook.application;
    application.load('calculationMode');
    return ctx.sync().then(function() {
        console.log(application.calculationMode);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

