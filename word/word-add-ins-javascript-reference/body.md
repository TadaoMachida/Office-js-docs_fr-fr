# Objet Body (interface API JavaScript pour Word)

Représente le corps d’un document ou d’une section.

_S’applique à : Word 2016, Word pour iPad, Word pour Mac_

## Propriétés
| Propriété   | Type|Description
|:---------------|:--------|:----------|
|style|string|Obtient ou définit le style utilisé pour le corps. Il s’agit du nom du style pré-installé ou personnalisé.|
|text|string|Obtient le texte du corps. Utilisez la méthode insertText pour insérer du texte. En lecture seule.|

_Voir des [exemples](#property-access-examples) d’accès aux propriétés._

## Relations
| Relation | Type|Description|
|:---------------|:--------|:----------|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Obtient la collection des objets de contrôle de contenu de texte enrichi qui se trouvent dans le corps. En lecture seule.|
|font|[Font](font.md)|Obtient le format de texte du corps. Utilisez cette propriété pour obtenir et définir le nom, la taille et la couleur de la police, ainsi que d’autres propriétés. En lecture seule.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Obtient la collection d’objets inlinePicture qui se trouvent dans le corps. La collection n’inclut pas d’images flottantes. En lecture seule.|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Obtient la collection d’objets de paragraphe qui se trouvent dans le corps. En lecture seule.|
|parentContentControl|[ContentControl](contentcontrol.md)|Obtient le contrôle de contenu qui contient le corps. Renvoie null s’il n’existe pas de contrôle de contenu parent. En lecture seule.|

## Méthodes

| Méthode   | Type renvoyé|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Efface le contenu de l’objet de corps. L’utilisateur peut effectuer l’opération d’annulation sur le contenu effacé.|
|[getHtml()](#gethtml)|string|Obtient la représentation HTML de l’objet de corps.|
|[getOoxml()](#getooxml)|string|Obtient la représentation OOXML (Office Open XML) de l’objet de corps.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Insère un saut à l’emplacement spécifié. Vous pouvez uniquement insérer un saut dans le corps de document principal, sauf s’il s’agit d’un saut de ligne, auquel cas il peut être inséré dans n’importe quel objet de corps. La valeur insertLocation peut être « Start » (début) ou « End » (fin).|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Encadre l’objet de corps avec un contrôle de contenu de texte enrichi.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Insère un document dans le corps à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Insère du code HTML à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertInlinePictureFromBase64base64EncodedImage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Insère une image dans le corps à l’emplacement spécifié. La valeur insertLocation peut être « Start » (début) ou « End » (fin). |
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Insère du code OOXML à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être « Start » (début) ou « End » (fin).|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Insère du texte dans le corps à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|Effectue une recherche avec les options de recherche spécifiées dans l’étendue de l’objet de corps. Les résultats de la recherche sont un ensemble d’objets de plage.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Sélectionne le corps et y accède via l’interface utilisateur de Word. Les valeurs selectionMode peuvent être « Select » (sélectionner), « Start » (début) ou « End » (fin).|

## Détails de méthodes

### clear()
Efface le contenu de l’objet de corps. L’utilisateur peut effectuer l’opération d’annulation sur le contenu effacé.

#### Syntaxe
```js
bodyObject.clear();
```

#### Paramètres
Aucun

#### Retourne
void

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to clear the contents of the body.
    body.clear();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Cleared the body contents.');
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

Le complément d’exemple [Silly stories](https://aka.ms/sillystorywordaddin) montre comment utiliser la méthode **clear** pour effacer le contenu d’un document.

### getHtml()
Obtient la représentation HTML de l’objet de corps.

#### Syntaxe
```js
bodyObject.getHtml();
```

#### Paramètres
Aucun

#### Retourne
string

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to get the HTML contents of the body.
    var bodyHTML = body.getHtml();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body HTML contents: " + bodyHTML.value);
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### getOoxml()
Obtient la représentation OOXML (Office Open XML) de l’objet de corps.

#### Syntaxe
```js
bodyObject.getOoxml();
```

#### Paramètres
Aucun

#### Retourne
string

#### Exemples 
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to get the OOXML contents of the body.
    var bodyOOXML = body.getOoxml();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body OOXML contents: " + bodyOOXML.value);
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Insère un saut à l’emplacement spécifié. Vous pouvez uniquement insérer un saut dans le corps de document principal, sauf s’il s’agit d’un saut de ligne, auquel cas il peut être inséré dans n’importe quel objet de corps. La valeur insertLocation peut être « Start » (début) ou « End » (fin).

#### Syntaxe
```js
bodyObject.insertBreak(breakType, insertLocation);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|breakType|BreakType|Obligatoire. Type de saut à ajouter au corps.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Start » (début) ou « End » (fin).|

#### Retourne
void

#### Détails supplémentaires
À l’exception des sauts de ligne, vous ne pouvez pas insérer de saut dans les en-têtes, les pieds de page, les notes de bas de page, les notes de fin, les commentaires et les zones de texte.  

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (ctx) {
    
    // Create a proxy object for the document body.
    var body = ctx.document.body;
    
    // Queue a commmand to insert a page break at the start of the document body.
    body.insertBreak(Word.BreakType.page, Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return ctx.sync().then(function () {
        console.log('Added a page break at the start of the document body.');
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### insertContentControl()
Encadre l’objet de corps avec un contrôle de contenu de texte enrichi.

#### Syntaxe
```js
bodyObject.insertContentControl();
```

#### Paramètres
Aucun

#### Retourne
[ContentControl](contentcontrol.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to wrap the body in a content control.
    body.insertContentControl();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Wrapped the body in a content control.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Insère un document dans le corps à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### Syntaxe
```js
bodyObject.insertFileFromBase64(base64File, insertLocation);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|base64File|string|Obligatoire. Contenu de fichier encodé au format Base64 à insérer.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### Retourne
[Range](range.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to insert base64 encoded .docx at the beginning of the content body.
    // You will need to implement getBase64() to pass in a string of a base64 encoded docx file.
    body.insertFileFromBase64(getBase64(), Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Added base64 encoded text to the beginning of the document body.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

Le complément d’exemple[Silly stories](https://aka.ms/sillystorywordaddin) montre comment utiliser la méthode **insertFileFromBase64** pour insérer des fichiers .docx à partir d’un service.

### insertHtml(html: string, insertLocation: InsertLocation)
Insère du code HTML à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### Syntaxe
```js
bodyObject.insertHtml(html, insertLocation);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|Html|string|Obligatoire. Code HTML à insérer dans le document.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### Retourne
[Range](range.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to insert HTML in to the beginning of the body.
    body.insertHtml('<strong>This is text inserted with body.insertHtml()</strong>', Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('HTML added to the beginning of the document body.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Insère une image dans le corps à l’emplacement spécifié. La valeur insertLocation peut être « Start » (début) ou « End » (fin).

#### Syntaxe
bodyObject.insertInlinePictureFromBase64(image, insertLocation);

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Obligatoire. Image encodée au format Base64 à insérer dans le corps.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Start » (début) ou « End » (fin).|

#### Renvoie
[InlinePicture](inlinepicture.md)

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Insère du code OOXML à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### Syntaxe
```js
bodyObject.insertOoxml(ooxml, insertLocation);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|ooxml|string|Obligatoire. Contenu OOXML ou wordProcessingML à insérer.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### Retourne
[Range](range.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to insert OOXML in to the beginning of the body.
    body.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('OOXML added to the beginning of the document body.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### Informations supplémentaires
Pour obtenir des instructions sur l’utilisation d’OOXML, voir [Création de compléments plus performants pour Word avec Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx). L’exemple [Word-Add-in-DocumentAssembly][body.insertOoxml] vous montre comment utiliser cette API pour assembler un document.

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être « Start » (début) ou « End » (fin).

#### Syntaxe
```js
bodyObject.insertParagraph(paragraphText, insertLocation);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|paragraphText|string|Obligatoire. Texte de paragraphe à insérer.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Start » (début) ou « End » (fin).|

#### Retourne
[Paragraph](paragraph.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to insert the paragraph at the end of the document body.
    body.insertParagraph('Content of a new paragraph', Word.InsertLocation.end);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Paragraph added at the end of the document body.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

#### Informations supplémentaires
L’exemple [Word-Add-in-DocumentAssembly][body.insertParagraph] vous montre comment utiliser la méthode insertParagraph pour assembler un document.

### insertText(text: string, insertLocation: InsertLocation)
Insère du texte dans le corps à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### Syntaxe
```js
bodyObject.insertText(text, insertLocation);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|texte|string|Obligatoire. Texte à insérer.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### Retourne
[Range](range.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to insert text in to the beginning of the body.
    body.insertText('This is text inserted with body.insertText()', Word.InsertLocation.start);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Text added to the beginning of the document body.');
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
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

#### Retourne
void

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      '; Font name: ' + body.font.name +
                      '; Font color: ' + body.font.color +
                      '; Body style: ' + body.style;

        console.log(results);
    });  
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```
### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Effectue une recherche avec les options de recherche spécifiées dans l’étendue de l’objet de corps. Les résultats de la recherche sont un ensemble d’objets de plage.

#### Syntaxe
```js
bodyObject.search(searchText, searchOptions);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|searchText|string|Obligatoire. Texte de recherche.|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|Facultatif. Options de la recherche.|

#### Renvoie
[SearchResultCollection](searchresultcollection.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to search the document.
    var searchResults = context.document.body.search('video', {matchCase: false});

    // Queue a commmand to load the results.
    context.load(searchResults, 'text, font');

    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        var results = 'Found count: ' + searchResults.items.length + 
                      '; we highlighted the results.';

        // Queue a command to change the font for each found item. 
        for (var i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].font.color = '#FF0000'    // Change color to Red
          searchResults.items[i].font.highlightColor = '#FFFF00';
          searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log(results);
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

#### Informations supplémentaires
L’exemple [Word-Add-in-DocumentAssembly][body.search] fournit un autre exemple de recherche d’un document.

### select(selectionMode: SelectionMode)
Sélectionne le corps et y accède via l’interface utilisateur de Word. Les valeurs selectionMode peuvent être « Select » (sélectionner), « Start » (début) ou « End » (fin).

#### Syntaxe
```js
bodyObject.select(selectionMode);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|selectionMode|SelectionMode|Facultatif. Le mode de sélection peut être « Select » (sélectionner), « Start » (début) ou « End » (fin). « Select » (sélectionner) est la valeur par défaut.|

#### Retourne
void

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to select the document body. The Word UI will 
    // move to the selected document body.
    body.select();
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Selected the document body.');
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

### Obtenir la propriété de texte sur l’objet de corps
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to load the text in document body.
    context.load(body, 'text');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });  
})
.catch(function (error) {
    console.log("Error: " + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### Obtenir le style et les propriétés de taille de police, nom de police et couleur de police sur l’objet de corps

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the document body.
    var body = context.document.body;
    
    // Queue a commmand to load font and style information for the document body.
    context.load(body, 'font/size, font/name, font/color, style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Show the results of the load method. Here we show the
        // property values on the body object.
        var results = 'Font size: ' + body.font.size +
                      '; Font name: ' + body.font.name +
                      '; Font color: ' + body.font.color +
                      '; Body style: ' + body.style;

        console.log(results);
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


[body.insertOoxml] : https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L127 "insert OOXML"[body.insertParagraph] : https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L153 "insert paragraph" [body.search] : https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L261 "body search"
