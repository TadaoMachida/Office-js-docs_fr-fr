# Objet ChartPoint (interface API JavaScript pour Excel)

Représente un point d’une série dans un graphique.

## Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|value|object|Renvoie la valeur d’un point du graphique. En lecture seule.|

## Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|format|[ChartPointFormat](chartpointformat.md)|Regroupe les propriétés de format d’un point d’un graphique. En lecture seule.|

## Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails des méthodes


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
