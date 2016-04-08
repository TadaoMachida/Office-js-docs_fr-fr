# Objet FilterDatetime (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Excel pour iOS, Office 2016_

Cet objet représente la méthode de filtrage d’une date lorsque des valeurs sont filtrées.

## Propriétés

| Propriété	   | Type	|Description
| :---------------| :--------| :----------||date|string|Date au format ISO8601 utilisée pour filtrer des données.||specificity|string|Spécificité de la date à utiliser pour conserver des données. Par exemple, si la date est 2005-04-02 et la spécificité est définie sur « mois », le filtre conservera toutes les lignes dont la date correspond au mois d’avril 2005. Les valeurs possibles sont les suivantes : Year, Monday, Day, Hour, Minute, Second.|_Consultez les [exemples](#property-access-examples) d’accès aux propriétés._

## Relations
Aucune


## Méthodes

| Méthode		   | Type de retour	|Description|| :---------------| :--------| :----------||[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails de la méthode


### load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### Syntaxe
```js
object.load(param);
```

#### Paramètres
| Paramètre	   | Type	|Description|| :---| :---| :---||param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### Renvoie
void

