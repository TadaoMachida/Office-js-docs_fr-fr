# Objet ChartAxisFormat (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Office 2016_

Regroupe les propriétés de format des axes du graphique.

## Propriétés

Aucun

## Relations
| Relation | Type|Description|
|:---------------|:--------|:----------|
|font|[ChartFont](chartfont.md)|Représente les attributs de police, comme le nom de la police, la taille de police, la couleur, etc. de l’élément d’axe du graphique. En lecture seule.|
|line|[ChartLineFormat](chartlineformat.md)|Représente le format des lignes du graphique. En lecture seule.|

## Méthodes

| Méthode   | Type renvoyé|Description|
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
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### Renvoie
void

