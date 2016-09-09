# Options de chargement d’objet (interface API JavaScript pour Excel)

Représente un objet qui peut être transmis à la méthode de chargement pour spécifier l’ensemble de propriétés et de relations à charger lors de l’exécution de la méthode sync() qui synchronise les états entre les objets Excel et les objets de proxy JavaScript correspondants dans le complément. Cet objet utilise des options telles que les paramètres des propriétés select et expand afin de spécifier l’ensemble de propriétés à charger sur l’objet et autorise la pagination sur la collection.

Vous pouvez également fournir une chaîne ou un tableau qui contient les propriétés et les relations à charger, tel qu’illustré dans l’exemple suivant.

```js   
object.load  ('<var1>,<relation1/var2>');

// Pass the parameter as an array.
object.load (["var1", "relation1/var2"]);
```

## Propriétés
| Propriété     | Type   |Description|
|:---------------|:--------|:----------|
|select|object|Fournissez un tableau ou une liste de noms de paramètres/relations (en les séparant par des virgules) à charger lors de l’appel de la méthode executeAsync. Par exemple, "propriété1, relation1", ["propriété1", "relation1"]. Facultatif.|
|expand|object|Fournissez un tableau ou une liste de noms de relations (en les séparant par des virgules) à charger lors de l’appel de la méthode executeAsync. Par exemple, "relation1, relation2", [ "relation1", "relation2"]. Facultatif.|
|top|int| Indiquez le nombre d’éléments de la collection demandée à inclure dans le résultat. Facultatif.|
|skip|int|Indiquez le nombre d’éléments de la collection devant être ignorés et exclus du résultat. Si une valeur est définie pour `top`, la sélection du résultat démarre une fois que le nombre spécifié d’éléments a été ignoré. Facultatif.|

#### Exemples

L’exemple sélectionne les 100 premières lignes du tableau.

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItem("Table1");
    var tableRows = table.rows.load({"select" : "index, values","top": 100, "skip": 0 })
    return ctx.sync().then(function() {
        for (var i = 0; i < tableRows.items.length; i++)
        {
            console.log(tableRows.items[i].index);
            console.log(tableRows.items[i].values);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
})
```
