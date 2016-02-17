# Objet Paragraph (interface API JavaScript pour Word)

Représente un seul paragraphe dans une sélection, une plage, un contrôle de contenu ou le corps d’un document.

_S’applique à : Word 2016, Word pour iPad, Word pour Mac_

## Propriétés
| Propriété   | Type|Description
|:---------------|:--------|:----------|
|outlineLevel|int|Obtient ou définit le niveau hiérarchique pour le paragraphe.|
|style|string|Obtient ou définit le style utilisé pour le paragraphe. Il s’agit du nom du style pré-installé ou personnalisé. L’exemple [Word-Add-in-DocumentAssembly][paragraph.style] vous montre comment définir le style de paragraphe.|
|texte|string|Obtient le texte du paragraphe. En lecture seule.|

_Voir des [exemples](#property-access-examples) d’accès aux propriétés._

## Relations
| Relation | Type|Description|
|:---------------|:--------|:----------|
|Alignment|**Alignment**|Obtient ou définit l’alignement d’un paragraphe. La valeur peut être « left » (gauche), « centered » (centré), « right » (droite) ou « justified » (justifié).|
|contentControls|[ContentControlCollection](contentcontrolcollection.md)|Obtient la collection d’objets de contrôle de contenu qui se trouvent dans le paragraphe. En lecture seule.|
|firstLineIndent|**float**|Renvoie ou définit la valeur, en points, du retrait de première ligne ou du retrait négatif. Utilisez une valeur positive pour définir un retrait de première ligne et une valeur négative pour définir un retrait négatif.|
|font|[Font](font.md)|Obtient le format de texte du paragraphe. Utilisez cette propriété pour obtenir et définir le nom de la police, la taille, la couleur et d’autres propriétés. En lecture seule.|
|inlinePictures|[InlinePictureCollection](inlinepicturecollection.md)|Obtient la collection d’objets inlinePicture qui se trouvent dans le paragraphe. La collection n’inclut pas d’images flottantes. En lecture seule.|
|leftIndent|**float**|Obtient ou définit la valeur de retrait à gauche, en points, pour le paragraphe.|
|lineSpacing|**float**|Obtient ou définit l’interligne, en points, pour le paragraphe spécifié. Dans l’interface utilisateur de Word, cette valeur est divisée par 12.|
|lineUnitAfter|**float**|Obtient ou définit l’espace, en lignes de quadrillage, après le paragraphe.|
|lineUnitBefore|**float**|Obtient ou définit la quantité d’espace, en lignes de quadrillage, avant le paragraphe.|
|parentContentControl|[ContentControl](contentcontrol.md)|Obtient le contrôle de contenu qui contient le paragraphe. Renvoie null s’il n’existe pas de contrôle de contenu parent. En lecture seule.|
|rightIndent|**float**|Obtient ou définit la valeur de retrait à droite, en points, pour le paragraphe.|
|spaceAfter|**float**|Obtient ou définit l’espacement, en points, après le paragraphe.|
|spaceBefore|**float**|Obtient ou définit l’espacement, en points, avant le paragraphe.|

## Méthodes

| Méthode   | Type renvoyé|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Efface le contenu de l’objet de paragraphe. L’utilisateur peut effectuer l’opération d’annulation sur le contenu effacé.|
|[delete()](#delete)|void|Supprime le paragraphe et son contenu du document.|
|[getHtml()](#gethtml)|string|Obtient la représentation HTML de l’objet de paragraphe.|
|[getOoxml()](#getooxml)|string|Obtient la représentation Office Open XML (OOXML) de l’objet de paragraphe.|
|[insertBreak(breakType: BreakType, insertLocation: InsertLocation)](#insertbreakbreaktype-breaktype-insertlocation-insertlocation)|void|Insère un saut à l’emplacement spécifié. Vous pouvez uniquement insérer un saut dans des paragraphes qui sont contenus dans le corps de document principal, sauf s’il s’agit d’un saut de ligne, auquel cas il peut être inséré dans n’importe quel objet de corps. La valeur insertLocation peut être « After » (après) ou « Before » (avant).|
|[insertContentControl()](#insertcontentcontrol)|[ContentControl](contentcontrol.md)|Encadre l’objet de paragraphe avec un contrôle de contenu de texte enrichi.|
|[insertFileFromBase64(base64File: string, insertLocation: InsertLocation)](#insertfilefrombase64base64file-string-insertlocation-insertlocation)|[Range](range.md)|Insère un document dans le paragraphe à l’emplacement spécifié. La valeur insertLocation peut être « Start » (début) ou « End » (fin).|
|[insertHtml(html: string, insertLocation: InsertLocation)](#inserthtmlhtml-string-insertlocation-insertlocation)|[Range](range.md)|Insère du code HTML dans le paragraphe à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)](#insertinlinepicturefrombase64base64encodedimage-string-insertlocation-insertlocation)|[InlinePicture](inlinepicture.md)|Insère une image dans le paragraphe à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant), « After » (après), « Start » (début) ou « End » (fin).|
|[insertOoxml(ooxml: string, insertLocation: InsertLocation)](#insertooxmlooxml-string-insertlocation-insertlocation)|[Range](range.md)|Insère du contenu OOXML ou wordProcessingML dans le paragraphe, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).|
|[insertText(text: string, insertLocation: InsertLocation)](#inserttexttext-string-insertlocation-insertlocation)|[Range](range.md)|Insère du texte dans le paragraphe à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|
|[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](#searchsearchtext-string-searchoptions-paramtypestringssearchoptions)|[SearchResultCollection](searchresultcollection.md)|Effectue une recherche avec les options de recherche spécifiées dans l’étendue de l’objet de paragraphe. Les résultats de la recherche sont un ensemble d’objets de plage.|
|[select(selectionMode: SelectionMode)](#selectselectionmode-selectionmode)|void|Sélectionne le paragraphe et y accède via l’interface utilisateur de Word. Le mode de sélection peut être « Select » (sélectionner), « Start » (début) ou « End » (fin). « Select » (sélectionner) est la valeur par défaut.|

## Détails de méthodes

### clear()
Efface le contenu de l’objet de paragraphe. L’utilisateur peut effectuer l’opération d’annulation sur le contenu effacé.

#### Syntaxe
```js
paragraphObject.clear();
```

#### Paramètres
Aucun

#### Retourne
void

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to clear the contents of the first paragraph.
        paragraphs.items[0].clear();    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Cleared the contents of the first paragraph.');
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

### delete()
Supprime le paragraphe et son contenu du document.

#### Syntaxe
```js
paragraphObject.delete();
```

#### Paramètres
Aucun

#### Retourne
void

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the text property for all of the paragraphs.
    context.load(paragraphs, 'text');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to delete the first paragraph.
        paragraphs.items[0].delete();    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Deleted the first paragraph.');
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

### getHtml()
Obtient la représentation HTML de l’objet de paragraphe.

#### Syntaxe
```js
paragraphObject.getHtml();
```

#### Paramètres
Aucun

#### Retourne
string

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a a set of commands to get the HTML of the first paragraph.
        var html = paragraphs.items[0].getHtml();    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph HTML: ' + html.value);
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

### getOoxml()
Obtient la représentation Office Open XML (OOXML) de l’objet de paragraphe.

#### Syntaxe
```js
paragraphObject.getOoxml();
```

#### Paramètres
Aucun

#### Retourne
string

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a a set of commands to get the OOXML of the first paragraph.
        var ooxml = paragraphs.items[0].getOoxml();    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Paragraph OOXML: ' + ooxml.value);
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

### insertBreak(breakType: BreakType, insertLocation: InsertLocation)
Insère un saut à l’emplacement spécifié. Vous pouvez uniquement insérer un saut dans des paragraphes qui sont contenus dans le corps de document principal, sauf s’il s’agit d’un saut de ligne, auquel cas il peut être inséré dans n’importe quel objet de corps. La valeur insertLocation peut être « Before » (avant) ou « After » (après).

#### Syntaxe
```js
paragraphObject.insertBreak(breakType, insertLocation);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|breakType|BreakType|Obligatoire. Type de saut à ajouter au document.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Before » (avant) » ou « After » (après).|

#### Retourne
void

#### Détails supplémentaires
Vous ne pouvez pas insérer de saut dans les en-têtes, les pieds de page, les notes de bas de page, les notes de fin, les commentaires et les zones de texte. 

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert a page break after the first paragraph.
        paragraph.insertBreak('page', 'After');    
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted a page break after the paragraph.');
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

### insertContentControl()
Encadre l’objet de paragraphe avec un contrôle de contenu de texte enrichi.

#### Syntaxe
```js
paragraphObject.insertContentControl();
```

#### Paramètres
Aucun

#### Retourne
[ContentControl](contentcontrol.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to wrap the first paragraph in a rich text content control.
        paragraph.insertContentControl(); 
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Wrapped the first paragraph in a content control.');
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
L’exemple [Word-Add-in-DocumentAssembly][paragraph.insertContentControl] vous montre comment utiliser la méthode insertContentControl.

### insertFileFromBase64(base64File: string, insertLocation: InsertLocation)
Insère un document dans le paragraphe à l’emplacement spécifié. La valeur insertLocation peut être « Start » (début) ou « End » (fin).

#### Syntaxe
```js
paragraphObject.insertFileFromBase64(base64File, insertLocation);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|base64File|string|Obligatoire. Contenu du fichier encodé au format Base64 à insérer.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Start » (début) ou « End » (fin).|

#### Retourne
[Range](range.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        // Queue a command to insert base64 encoded .docx at the beginning of the first paragraph.
        // This won't work unless you have a definition for getBase64().
        paragraph.insertFileFromBase64(getBase64(), Word.InsertLocation.start);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted base64 encoded content at the beginning of the first paragraph.');
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

### insertHtml(html: string, insertLocation: InsertLocation)
Insère du code HTML dans le paragraphe à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### Syntaxe
```js
paragraphObject.insertHtml(html, insertLocation);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|Html|string|Obligatoire. Code HTML à insérer dans le paragraphe.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### Retourne
[Range](range.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert HTML content at the end of the first paragraph.
        paragraph.insertHtml('<strong>Inserted HTML.</strong>', Word.InsertLocation.end);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted HTML content at the end of the first paragraph.');
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

### insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)
Insère une image dans le paragraphe à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant), « After » (après), « Start » (début) ou « End » (fin).

#### Syntaxe
```js
paragraphObject.insertInlinePictureFromBase64(base64EncodedImage, insertLocation);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|base64EncodedImage|string|Obligatoire. Code HTML à insérer dans le paragraphe.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être définie sur « Before » (avant), « After » (après), « Start » (début) ou « End » (fin).|

#### Renvoie
[InlinePicture](inlinepicture.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];

        var b64encodedImg = "iVBORw0KGgoAAAANSUhEUgAAAB4AAAANCAIAAAAxEEnAAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACFSURBVDhPtY1BEoQwDMP6/0+XgIMTBAeYoTqso9Rkx1zG+tNj1H94jgGzeNSjteO5vtQQuG2seO0av8LzGbe3anzRoJ4ybm/VeKEerAEbAUpW4aWQCmrGFWykRzGBCnYy2ha3oAIq2MloW9yCCqhgJ6NtcQsqoIKdjLbFLaiACnYyf2fODbrjZcXfr2F4AAAAAElFTkSuQmCC";

        // Queue a command to insert a base64 encoded image at the beginning of the first paragraph.
        paragraph.insertInlinePictureFromBase64(b64encodedImg, Word.InsertLocation.start);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Added an image to the first paragraph.');
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
L’exemple [Word-Add-in-DocumentAssembly][paragraph.insertpicture] fournit un autre exemple de la façon d’insérer une image dans un paragraphe.

### insertOoxml(ooxml: string, insertLocation: InsertLocation)
Insère du contenu OOXML ou wordProcessingML dans le paragraphe, à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### Syntaxe
```js
paragraphObject.insertOoxml(ooxml, insertLocation);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|ooxml|string|Obligatoire. Contenu OOXML ou wordProcessingML à insérer dans le paragraphe.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).|

#### Retourne
[Range](range.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert Ooxml content into the first paragraph.
        var ooxmlContent = "<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
        paragraph.insertOoxml(ooxmlContent, Word.InsertLocation.end);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted OOXML at the end of the first paragraph.');
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
Pour obtenir des instructions sur l'utilisation d’OOXML, voir [Création de compléments plus performants pour Word avec Office Open XML](https://msdn.microsoft.com/en-us/library/office/dn423225.aspx).

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Insère un paragraphe à l’emplacement spécifié. La valeur insertLocation peut être définie sur « Before » (avant) ou « After » (après).

#### Syntaxe
```js
paragraphObject.insertParagraph(paragraphText, insertLocation);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|paragraphText|string|Obligatoire. Texte de paragraphe à insérer.|
|insertLocation|InsertLocation|Obligatoire. La valeur peut être « Before » (avant) » ou « After » (après).|

#### Retourne
[Paragraph](paragraph.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert the paragraph after the current paragraph.
        paragraph.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted a new paragraph at the end of the first paragraph.');
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

### insertText(text: string, insertLocation: InsertLocation)
Insère du texte dans le paragraphe à l’emplacement spécifié. La valeur insertLocation peut être « Replace » (remplacer), « Start » (début) ou « End » (fin).

#### Syntaxe
```js
paragraphObject.insertText(text, insertLocation);
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
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to insert text into the end of the paragraph.
        paragraph.insertText('New text inserted into the paragraph.', Word.InsertLocation.end);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Inserted text at the end of the first paragraph.');
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
    
    // Queue a commmand to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    context.load(paragraphs, {select: 'style', top: 2} );
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the first paragraph.
        var paragraph = paragraphs.items[0];        
        
        // Queue a command to load font information for the paragraph.
        context.load(paragraph, 'font/size, font/name, font/color');
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            // Show the results of the load method. Here we show the
            // property values on the paragraph object. Note that we 
            // requested the style property in the first load command.
            var results = "<strong>Paragraph</strong>--" +
                          "--Font size: " + paragraph.font.size +
                          "--Font name: " + paragraph.font.name +
                          "--Font color: " + paragraph.font.color +
                          "--Style: " + paragraph.style;

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

### search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)
Effectue une recherche avec les options de recherche spécifiées dans l’étendue de l’objet de paragraphe. Les résultats de la recherche sont un ensemble d’objets de plage.

#### Syntaxe
```js
paragraphObject.search(searchText, searchOptions);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|searchText|string|Obligatoire. Texte de recherche.|
|[searchOptions](searchoptions.md)|ParamTypeStrings.SearchOptions|Facultatif. Options de la recherche.|

#### Renvoie
[SearchResultCollection](searchresultcollection.md)

### select(selectionMode: SelectionMode)
Sélectionne le paragraphe et y accède via l’interface utilisateur de Word.

#### Syntaxe
```js
paragraphObject.select(selectionMode);
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
    
    // Create a proxy object for the paragraphs collection.
    var paragraphs = context.document.body.paragraphs;
    
    // Queue a commmand to load the style property for all of the paragraphs.
    context.load(paragraphs, 'style');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        
        // Queue a command to get the last paragraph a create a 
        // proxy paragraph object.
        var paragraph = paragraphs.items[paragraphs.items.length - 1]; 
        
        // Queue a command to select the paragraph. The Word UI will 
        // move to the selected paragraph.
        paragraph.select();
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Selected the last paragraph.');
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

## Informations de prise en charge

Utilisez l’[ensemble de conditions requises](https://msdn.microsoft.com/EN-US/library/office/mt590206.aspx) dans les vérifications à l’exécution pour vous assurer que votre application est prise en charge par la version d’hôte de Word. Pour plus d’informations sur la configuration requise pour le serveur et l’application d’hôte Office, voir [Configuration requise pour exécuter des compléments Office](https://msdn.microsoft.com/EN-US/library/office/dn833104.aspx). 


[paragraph.insertContentControl] : https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L161 "insert content control"[paragraph.style] : https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L172 "set style" [paragraph.insertpicture] : https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L236 "insert picture"
