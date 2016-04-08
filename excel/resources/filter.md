# Objet Filter (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Excel pour iOS, Office 2016_

Gère le filtrage de la colonne d’un tableau.

## Propriétés

Aucun

## Relations
| Relation | Type|Description|
|:---------------|:--------|:----------|
|de réussite|[FilterCriteria](filtercriteria.md)|Le filtre actuellement appliqué à la colonne donnée. En lecture seule.|

## Méthodes

| Méthode   | Type renvoyé|Description|
|:---------------|:--------|:----------|
|[apply(criteria: FilterCriteria)](#applycriteria-filtercriteria)|void|Appliquer les critères de filtre donnés à la colonne indiquée. La même fonctionnalité peut être obtenue avec l’une des méthodes d’assistance suivantes.|
|[applyBottomItemsFilter(count: number)](#applybottomitemsfiltercount-number)|void|Appliquer un filtre « Élément inférieur » à la colonne pour le nombre d’éléments donné.|
|[applyBottomPercentFilter(percent: number)](#applybottompercentfilterpercent-number)|void|Appliquer un filtre « Pourcentage inférieur » à la colonne pour le pourcentage d’éléments donné.|
|[applyCellColorFilter(color: string)](#applycellcolorfiltercolor-string)|void|Appliquer un filtre « Couleur de cellule » à la colonne pour la couleur donnée.|
|[applyCustomFilter(criteria1: string, criteria2: string, oper: FilterOperator)](#applycustomfiltercriteria1-string-criteria2-string-oper-filteroperator)|void|Appliquer un filtre « Icône » à la colonne pour les chaînes de critères données.|
|[applyDynamicFilter(criteria: string)](#applydynamicfiltercriteria-string)|void|Appliquer un filtre « Dynamique » à la colonne.|
|[applyFontColorFilter(color: string)](#applyfontcolorfiltercolor-string)|void|Appliquer un filtre « Couleur de police » à la colonne pour la couleur donnée.|
|[applyIconFilter(icon: Icon)](#applyiconfiltericon-icon)|void|Appliquer un filtre « Icône » à la colonne pour l’icône donnée.|
|[applyTopItemsFilter(count: number)](#applytopitemsfiltercount-number)|void|Appliquer un filtre « Élément supérieur » à la colonne pour le nombre d’éléments donné.|
|[applyTopPercentFilter(percent: number)](#applytoppercentfilterpercent-number)|void|Appliquer un filtre « Pourcentage supérieur » à la colonne pour le pourcentage d’éléments donné.|
|[applyValuesFilter(values: ()[])](#applyvaluesfiltervalues-)|void|Appliquer un filtre « Valeurs » à la colonne pour les valeurs données.|
|[clear()](#clear)|void|Effacer le filtre sur la colonne donnée.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails des méthodes


### apply(criteria: FilterCriteria)
Appliquer les critères de filtre donnés à la colonne indiquée. La même fonctionnalité peut être obtenue avec l’une des méthodes d’assistance suivantes. 

#### Syntaxe
```js
filterObject.apply(criteria);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|de réussite|FilterCriteria|Critères à appliquer.|

#### Renvoie
void

#### Exemple
L’exemple suivant indique comment appliquer un filtre personnalisé avec la méthode apply() générique.

```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    var filterCriteria = { 
		filterOn: Excel.FilterOn.custom,
		criterion1: ">50",
		operator: Excel.FilterOperator.and,
		criterion2: "<100"
    	} 
    column.filter.apply(filterCriteria);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyBottomItemsFilter(count: number)
Appliquer un filtre « Élément inférieur » à la colonne pour le nombre d’éléments donné.

#### Syntaxe
```js
filterObject.applyBottomItemsFilter(count);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|count|number|Nombre d’éléments à partir du bas à afficher.|

#### Renvoie
void

#### Exemple
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyBottomPercentFilter(percent: number)
Appliquer un filtre « Pourcentage inférieur » à la colonne pour le pourcentage d’éléments donné.

#### Syntaxe
```js
filterObject.applyBottomPercentFilter(percent);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|pourcentage|number|Pourcentage d’éléments à partir du bas à afficher.|

#### Renvoie
void

#### Exemple
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyBottomPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### applyCellColorFilter(color: string)
Appliquer un filtre « Couleur de cellule » à la colonne pour la couleur donnée.


#### Syntaxe
```js
filterObject.applyCellColorFilter(color);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|color|string|Couleur d’arrière-plan des cellules à afficher.|

#### Renvoie
void

#### Exemple
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCellColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyCustomFilter(criteria1: string, criteria2: string, oper: FilterOperator)
Appliquer un filtre « Icône » à la colonne pour les chaînes de critères données.

#### Syntaxe
```js
filterObject.applyCustomFilter(criteria1, criteria2, oper);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|criteria1|string|Première chaîne de critères.|
|criteria2|string|Facultatif. Deuxième chaîne de critères.|
|oper|FilterOperator|Facultatif. Opérateur qui décrit comment les deux critères sont joints.|

#### Renvoie
void


#### Exemple
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyCustomFilter('>50','<100','and');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyDynamicFilter(criteria: string)
Appliquer un filtre « Dynamique » à la colonne.

#### Syntaxe
```js
filterObject.applyDynamicFilter(criteria);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|de réussite|string|Critères dynamiques à appliquer.  Les valeurs possibles sont les suivantes : Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday|

#### Renvoie
void

#### Exemple
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyDynamicFilter(Excel.DynamicFilterCriteria.aboveAverage);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyFontColorFilter(color: string)
Appliquer un filtre « Couleur de police » à la colonne pour la couleur donnée.

#### Syntaxe
```js
filterObject.applyFontColorFilter(color);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|color|string|Couleur de police des cellules à afficher.|

#### Renvoie
void

#### Exemple
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyFontColorFilter('red');
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyIconFilter(icon: Icon)
Appliquer un filtre « Icône » à la colonne pour l’icône donnée.

#### Syntaxe
```js
filterObject.applyIconFilter(icon);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|icône|Icône|Icônes des cellules à afficher.|

#### Renvoie
void

#### Exemple
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyIconFilter(Excel.icons.fiveArrows.yellowDownInclineArrow);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### applyTopItemsFilter(count: number)
Appliquer un filtre « Élément supérieur » à la colonne pour le nombre d’éléments donné.

#### Syntaxe
```js
filterObject.applyTopItemsFilter(count);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|count|number|Nombre d’éléments à partir du haut à afficher.|

#### Renvoie
void

#### Exemple
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopItemsFilter(3);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### applyTopPercentFilter(percent: number)
Appliquer un filtre « Pourcentage supérieur » à la colonne pour le pourcentage d’éléments donné.

#### Syntaxe
```js
filterObject.applyTopPercentFilter(percent);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|pourcentage|number|Pourcentage d’éléments à partir du haut à afficher.|

#### Renvoie
void

#### Exemple
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyTopPercentFilter(30);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
### applyValuesFilter(values: ()[])
Appliquer un filtre « Valeurs » à la colonne pour les valeurs données.

#### Syntaxe
```js
filterObject.applyValuesFilter(values);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|values|()[]|Liste des valeurs à afficher.|

#### Renvoie
void

#### Exemple
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.applyValuesFilter(['a','b']);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### clear()
Effacer le filtre sur la colonne donnée.

#### Syntaxe
```js
filterObject.clear();
```

#### Paramètres
Aucun

#### Retourne
void

#### Exemple
```js
Excel.run(function (ctx) { 
    var column = ctx.workbook.tables.getItem("Table1").columns.getItemAt(0);
    column.filter.clear();
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
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### Renvoie
void

