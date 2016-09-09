# Objet ContentControlCollection (interface API JavaScript pour Word)

Contient une collection d’objets ContentControl. Les contrôles de contenu sont des régions liées et potentiellement étiquetées d’un document qui servent de conteneur pour des types de contenu spécifiques. Chaque contrôle de contenu peut comporter des images, des tableaux ou des paragraphes de texte mis en forme. Actuellement, seuls les contrôles de contenu à texte enrichi sont pris en charge.

_S’applique à : Word 2016, Word pour iPad, Word pour Mac_

## Propriétés
| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|Items|[ContentControl[]](contentcontrol.md)|Collection d’objets contentControl. En lecture seule.|

## Relations
Aucun


## Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[getById(id: number)](#getbyidid-number)|[ContentControl](contentcontrol.md)|Obtient un contrôle de contenu par son identificateur.|
|[getByTag(tag: string)](#getbytagtag-string)|[ContentControlCollection](contentcontrolcollection.md)|Obtient les contrôles de contenu qui portent l’indicateur spécifié.|
|[getByTitle(title: string)](#getbytitletitle-string)|[ContentControlCollection](contentcontrolcollection.md)|Obtient les contrôles de contenu qui ont le titre spécifié.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails de méthodes

### getById(id: number)
Obtient un contrôle de contenu par son identificateur.

#### Syntaxe
```js
contentControlCollectionObject.getById(id);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|id|number|Obligatoire. Identificateur de contrôle de contenu.|

#### Retourne
[ContentControl](contentcontrol.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content control that contains a specific id.
    var contentControl = context.document.contentControls.getById(30086310);

    // Queue a command to load the text property for a content control.
    context.load(contentControl, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The content control with that Id has been found in this document.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

### getByTag(tag: string)
Obtient les contrôles de contenu qui portent l’indicateur spécifié.

#### Syntaxe
```js
contentControlCollectionObject.getByTag(tag);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|tag|string|Obligatoire. Indicateur défini sur un contrôle de contenu.|

#### Retourne
[ContentControlCollection](contentcontrolcollection.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection that contains a specific tag.
    var contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');

    // Queue a command to load the text property for all of content controls with a specific tag.
    context.load(contentControlsWithTag, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTag.items.length === 0) {
            console.log("There isn't a content control with a tag of Customer-Address in this document.");
        } else {
            console.log('The first content control with the tag of Customer-Address has this text: ' + contentControlsWithTag.items[0].text);
        }

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
L’exemple [Word-Add-in-DocumentAssembly][contentControls.getByTag] est un autre exemple d’utilisation de la méthode getByTag.


### getByTitle(title: string)
Obtient les contrôles de contenu qui ont le titre spécifié.

#### Syntaxe
```js
contentControlCollectionObject.getByTitle(title);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|title|string|Obligatoire. Titre d’un contrôle de contenu.|

#### Retourne
[ContentControlCollection](contentcontrolcollection.md)

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection that contains a specific title.
    var contentControlsWithTitle = context.document.contentControls.getByTitle('Enter Customer Address Here');

    // Queue a command to load the text property for all of content controls with a specific title.
    context.load(contentControlsWithTitle, 'text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControlsWithTitle.items.length === 0) {
            console.log("There isn't a content control with a title of 'Enter Customer Address Here' in this document.");
        } else {
            console.log("The first content control with the title of 'Enter Customer Address Here' has this text: " + contentControlsWithTitle.items[0].text);
        }

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
L’exemple [Word-Add-in-DocumentAssembly][contentControls.getByTitle] est un autre exemple d’utilisation de la méthode getByTitle.

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

#### Exemples
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;

    // Queue a command to load the id property for all of the content controls.
    context.load(contentControls, 'id');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        if (contentControls.items.length === 0) {
            console.log('No content control found.');
        }
        else {
            // Queue a command to load the properties on the first content control.
            contentControls.items[0].load(  'appearance,' +
                                            'cannotDelete,' +
                                            'cannotEdit,' +
                                            'id,' +
                                            'placeHolderText,' +
                                            'removeWhenEdited,' +
                                            'title,' +
                                            'text,' +
                                            'type,' +
                                            'style,' +
                                            'tag,' +
                                            'font/size,' +
                                            'font/name,' +
                                            'font/color');

            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    console.log('Property values of the first content control:' +
                        '   ----- appearance: ' + contentControls.items[0].appearance +
                        '   ----- cannotDelete: ' + contentControls.items[0].cannotDelete +
                        '   ----- cannotEdit: ' + contentControls.items[0].cannotEdit +
                        '   ----- color: ' + contentControls.items[0].color +
                        '   ----- id: ' + contentControls.items[0].id +
                        '   ----- placeHolderText: ' + contentControls.items[0].placeholderText +
                        '   ----- removeWhenEdited: ' + contentControls.items[0].removeWhenEdited +
                        '   ----- title: ' + contentControls.items[0].title +
                        '   ----- text: ' + contentControls.items[0].text +
                        '   ----- type: ' + contentControls.items[0].type +
                        '   ----- style: ' + contentControls.items[0].style +
                        '   ----- tag: ' + contentControls.items[0].tag +
                        '   ----- font size: ' + contentControls.items[0].font.size +
                        '   ----- font name: ' + contentControls.items[0].font.name +
                        '   ----- font color: ' + contentControls.items[0].font.color);
            });
        }
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

Le complément d’exemple [Silly stories](https://aka.ms/sillystorywordaddin) montre comment utiliser la méthode **load** pour charger la collection de contrôles de contenu avec les propriétés **tag** et **title**.

## Informations de prise en charge
Utilisez l’[ensemble de conditions requises](../office-add-in-requirement-sets.md) dans les vérifications à l’exécution pour vous assurer que votre application est prise en charge par la version d’hôte de Word. Pour plus d’informations sur la configuration requise pour le serveur et l’application d’hôte Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).


[contentControls.getByTag]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L300 "get by tag"
[contentControls.getByTitle]: https://github.com/OfficeDev/Word-Add-in-DocumentAssembly/blob/master/WordAPIDocAssemblySampleWeb/App/Home/Home.js#L331 "get by title"

