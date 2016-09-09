# Objet Table (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Représente une table dans une page OneNote.

## Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|borderVisible|bool|Obtient ou définit si les bordures sont visibles ou non. True si elles sont visibles, false si elles sont masquées.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-borderVisible)|
|columnCount|int|Obtient le nombre de colonnes dans le tableau. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-columnCount)|
|id|chaîne|Obtient l’ID du tableau. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-id)|
|rowCount|int|Obtient le nombre de lignes dans le tableau. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-rowCount)|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
| Relation | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|paragraph|[Paragraph](paragraph.md)|Obtient l’objet Paragraph qui contient l’objet Table. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-paragraph)|
|Objet Rows|[TableRowCollection](tablerowcollection.md)|Obtient toutes les lignes de la table. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-rows)|

## Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[appendColumn(values: string[])](#appendcolumnvalues-string)|void|Ajoute une colonne à la fin du tableau. Si elles sont spécifiées, les valeurs sont définies dans la nouvelle colonne. Dans le cas contraire, la colonne est vide.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-appendColumn)|
|[appendRow(values: string[])](#appendrowvalues-string)|[TableRow](tablerow.md)|Ajoute une ligne à la fin du tableau. Si elles sont spécifiées, les valeurs sont définies dans la nouvelle ligne. Dans le cas contraire, la ligne est vide.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-appendRow)|
|[clear()](#clear)|void|Efface le contenu du tableau.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-clear)|
|[getCell(rowIndex: number, cellIndex: number)](#getcellrowindex-number-cellindex-number)|[TableCell](tablecell.md)|Obtient la cellule du tableau correspondant à une ligne et une colonne spécifiées.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-getCell)|
|[insertColumn(index: number, values: string[])](#insertcolumnindex-number-values-string)|void|Insère une colonne au niveau de l’index donné dans le tableau. Si elles sont spécifiées, les valeurs sont définies dans la nouvelle colonne. Dans le cas contraire, la colonne est vide.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-insertColumn)|
|[insertRow(index: number, values: string[])](#insertrowindex-number-values-string)|[TableRow](tablerow.md)|Insère une ligne à l’index donné dans le tableau. Si elles sont spécifiées, les valeurs sont définies dans la nouvelle ligne. Dans le cas contraire, la ligne est vide.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-insertRow)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-load)|
|[setShadingColor(colorCode: string)](#setshadingcolorcolorcode-string)|void|Définit la couleur d’ombrage de toutes les cellules du tableau.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-table-setShadingColor)|

## Détails des méthodes


### appendColumn(values: string[])
Ajoute une colonne à la fin du tableau. Si elles sont spécifiées, les valeurs sont définies dans la nouvelle colonne. Dans le cas contraire, la colonne est vide.

#### Syntaxe
```js
tableObject.appendColumn(values);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|values|string[]|Facultatif. Facultatif. Chaînes à insérer dans la nouvelle colonne, spécifiées sous forme de tableau. Elles ne doivent pas contenir plus de valeurs que de lignes dans le tableau.|

#### Retourne
void

#### Exemples
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, append a column.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                table.appendColumn(["cell0", "cell1"]);
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### appendRow(values: string[])
Ajoute une ligne à la fin du tableau. Si elles sont spécifiées, les valeurs sont définies dans la nouvelle ligne. Dans le cas contraire, la ligne est vide.

#### Syntaxe
```js
tableObject.appendRow(values);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|values|string[]|Facultatif. Facultatif. Chaînes à insérer dans la nouvelle ligne, spécifiées sous forme de tableau. Elles ne doivent pas contenir plus de valeurs que de colonnes dans le tableau.|

#### Retourne
[TableRow](tablerow.md)

#### Exemples
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, append a column.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var row = table.appendRow(["cell0", "cell1"]);
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### clear()
Efface le contenu du tableau.

#### Syntaxe
```js
tableObject.clear();
```

#### Paramètres
Aucun

#### Retourne
void

### getCell(rowIndex: number, cellIndex: number)
Obtient la cellule du tableau correspondant à une ligne et une colonne spécifiées.

#### Syntaxe
```js
tableObject.getCell(rowIndex, cellIndex);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|rowIndex|number|Index de la ligne.|
|cellIndex|number|Index de la cellule dans la ligne.|

#### Retourne
[TableCell](tablecell.md)

#### Exemples
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, get a cell in the second row and third column.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var cell = table.getCell(2 /*Row Index*/, 3 /*Column Index*/);
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### insertColumn(index: number, values: string[])
Insère une colonne au niveau de l’index donné dans le tableau. Si elles sont spécifiées, les valeurs sont définies dans la nouvelle colonne. Dans le cas contraire, la colonne est vide.

#### Syntaxe
```js
tableObject.insertColumn(index, values);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Index au niveau duquel la colonne est insérée dans le tableau.|
|values|string[]|Facultatif. Facultatif. Chaînes à insérer dans la nouvelle colonne, spécifiées sous forme de tableau. Elles ne doivent pas contenir plus de valeurs que de lignes dans le tableau.|

#### Retourne
void

#### Exemples
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, insert a column at index two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                table.insertColumn(2, ["cell0", "cell1"]);
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### insertRow(index: number, values: string[])
Insère une ligne à l’index donné dans le tableau. Si elles sont spécifiées, les valeurs sont définies dans la nouvelle ligne. Dans le cas contraire, la ligne est vide.

#### Syntaxe
```js
tableObject.insertRow(index, values);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Index au niveau duquel la ligne est insérée dans le tableau.|
|values|string[]|Facultatif. Facultatif. Chaînes à insérer dans la nouvelle ligne, spécifiées sous forme de tableau. Elles ne doivent pas contenir plus de valeurs que de colonnes dans le tableau.|

#### Retourne
[TableRow](tablerow.md)

#### Exemples
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, insert a row at index two.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                var row = table.insertRow(2, ["cell0", "cell1"]);
            }
        }
        return ctx.sync();
    })
})
.catch(function(error) {
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

#### Retourne
void

### setShadingColor(colorCode: string)
Définit la couleur d’ombrage de toutes les cellules du tableau.

#### Syntaxe
```js
tableObject.setShadingColor(colorCode);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|colorCode|chaîne|Code de couleur selon lequel sont définies les cellules/param|

#### Retourne
void
### Exemples d’accès aux propriétés
**columnCount, rowCount, id**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // For each table, log properties.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                ctx.load(table);
                return ctx.sync().then(function() {
                    console.log("Table Id: " + table.id);
                    console.log("Row Count: " + table.rowCount);
                    console.log("Column Count: " + table.columnCount);
                    return ctx.sync();
                });
            }
        }
    });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**paragraph, rows**
```js
OneNote.run(function(ctx) {
    var app = ctx.application;
    var outline = app.getActiveOutline();
    
    // Queue a command to load outline.paragraphs and their types.
    ctx.load(outline, "paragraphs, paragraphs/type");
    
    // Run the queued commands, and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        var paragraphs = outline.paragraphs;
        
        // for each table, log its paragraph id.
        for (var i = 0; i < paragraphs.items.length; i++) {
            var paragraph = paragraphs.items[i];
            if (paragraph.type == "Table") {
                var table = paragraph.table;
                ctx.load(table, "paragraph/id, rows/id");
                return ctx.sync().then(function() {
                    console.log("Paragraph Id: " + table.paragraph.id);
                    var rows = table.rows;
                    
                    // for each rows in the table, log row index and id.
                    for (var i = 0; i < rows.items.length; i++) {
                        console.log("Row " + i + " Id: " + rows.items[i].id);
                    }
                    return ctx.sync();
                });
            }
        }
    })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

