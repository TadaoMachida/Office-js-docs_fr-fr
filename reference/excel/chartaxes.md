# Objet ChartAxes (interface API JavaScript pour Excel)

Représente les axes du graphique.

## Propriétés

Aucun

## Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|categoryAxis|[ChartAxis](chartaxis.md)|Représente l’axe des abscisses d’un graphique. En lecture seule.|
|seriesAxis|[ChartAxis](chartaxis.md)|Représente l’axe de séries d’un graphique 3D. En lecture seule.|
|valueAxis|[ChartAxis](chartaxis.md)|Représente l’axe des ordonnées. En lecture seule.|

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
