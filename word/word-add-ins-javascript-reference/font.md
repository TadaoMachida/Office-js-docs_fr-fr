# Objet Font (interface API JavaScript pour Word)

Représente une police.

_S’applique à : Word 2016, Word pour iPad, Word pour Mac_

## Propriétés
| Propriété   | Type|Description
|:---------------|:--------|:----------|
|bold|bool|Obtient ou définit une valeur qui indique si la police en gras. Renvoie true si la police est mise en forme en gras, sinon, false.|
|color|string|Obtient ou définit la couleur de la police spécifiée. Vous pouvez fournir la valeur au format « #RRGGBB » ou avec le nom de la couleur.|
|doubleStrikeThrough|bool|Obtient ou définit une valeur qui indique si la police est barrée double. Renvoie true si la police est mise en forme en tant que texte barré double, sinon, false.|
|highlightColor|string|Obtient ou définit la couleur de mise en surbrillance pour la police spécifiée. Vous pouvez fournir la valeur au format « #RRGGBB » ou avec le nom de la couleur.|
|italic|bool|Obtient ou définit une valeur qui indique si la police est en italique. Renvoie true si la police est en italique, sinon, false.|
|name|string|Obtient ou définit une valeur qui représente le nom de la police.|
|strikeThrough|bool|Obtient ou définit une valeur qui indique si la police est barrée. Renvoie true si la police est mise en forme en tant que texte barré, sinon, false.|
|subscript|bool|Obtient ou définit une valeur qui indique si la police correspond à du texte mis en indice. Renvoie true si la police correspond à du texte mis en indice, sinon, false.|
|superscript|bool|Obtient ou définit une valeur qui indique si la police correspond à du texte en exposant. Renvoie true si la police correspond à du texte mis en exposant, sinon, false.|

_Voir des [exemples](#property-access-examples) d’accès aux propriétés._

## Relations
| Relation | Type|Description|
|:---------------|:--------|:----------|
|size|**float**|Obtient ou définit une valeur qui représente la taille de police en points.|
|Souligné|[UnderlineType](underlinetype.md)|Obtient ou définit une valeur qui indique le type de trait de soulignement de la police. Les valeurs valides sont : « None » (aucun), « Single » (simple), « Word » (mot), « Double » (double), « Dotted » (pointillés), « Hidden » (masqué), « Thick » (épais), « Dashline » (tirets), « Dotline » (points), « DotDashLine » (ligne point-tiret), « TwoDotDashLine » (ligne point-point-tiret) et « Wave » (ondulé).|

## Méthodes

| Méthode   | Type renvoyé|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails de méthodes

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

#### Retourne
void

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the font property for all of the paragraphs.
    context.load(paragraphs, 'font');

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Create a proxy object for the font object on the first paragraph in the collection.
        var font = paragraphs.items[0].font;
        
        // Queue a set of property value changes on the font proxy object.
        font.size = 32;
        font.bold = true;
        font.color = '#0000ff';
        font.highlightColor = '#ffff00';
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('The font has changed.');
        });  
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## Exemples d’accès aux propriétés

### Modifier le nom de la police
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to change the current selection's font name.
    selection.font.name = 'Arial';
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font name has changed.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Modifier la couleur de la police
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to change the font color of the current selection.
    selection.font.color = 'blue'; 
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font color of the selection has been changed.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Modifier la taille de la police
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to change the current selection's font size.
    selection.font.size = 20;
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The font size has changed.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Mettre le texte sélectionné en surbrillance
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to highlight the current selection.
    selection.font.highlightColor = '#FFFF00'; // Yellow
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection has been highlighted.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Texte en gras
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to make the current selection bold.
    selection.font.bold = true;
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection is now bold.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});

```

### Texte souligné
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to underline the current selection.
    selection.font.underline = Word.UnderlineType.thick;
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection now has an underline style.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### Texte barré
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();
    
    // Queue a commmand to strikethrough the font of the current selection.
    selection.font.strikeThrough = true; 
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection now has a strikethrough.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## Informations de prise en charge

Utilisez l’[ensemble de conditions requises](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) dans les vérifications à l’exécution pour vous assurer que votre application est prise en charge par la version d’hôte de Word. Pour plus d’informations sur la configuration requise pour le serveur et l’application d’hôte Office, voir [Configuration requise pour exécuter des compléments Office](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx). 
