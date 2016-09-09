# Objet RequestContext (interface API JavaScript pour Word)

L’objet RequestContext facilite les demandes du complément auprès de l’application Word (rappelez-vous que les deux applications utilisent des processus différents).

_S’applique à : Word 2016, Word pour iPad, Word pour Mac_

## Propriétés
Aucun

## Méthodes

| Méthode         | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[load(object: object, option: object)](#loadobject-object-option-object)  |void     |Insère l’objet de proxy créé dans le calque JavaScript avec les propriétés et les options spécifiées dans le paramètre.|
|[sync()](#sync)  |Objet de promesse |Envoie les demandes en file d’attente à Word et renvoie un objet de promesse, qui peut être utilisé pour ajouter d’autres actions en chaîne.|

## Détails de méthodes

### load(object: object, option: object)
Insère l’objet de proxy créé dans le calque JavaScript avec les propriétés et les options spécifiées dans le paramètre.

#### Syntaxe
```js
requestContextObject.load(object, loadOption);
```

#### Paramètres
| Paramètre       | Type    |Description|
|:----------------|:--------|:----------|
|object|object|Facultatif. Indiquez le nom de l’objet à charger.|
|Option|[loadOption](loadoption.md)|Propriété facultative, mais recommandée. Spécifiez les options de chargement (select, expand, skip ou top). |

#### Retourne
void

##### Exemples

L’exemple suivant montre comment charger la propriété de texte sur une collection de paragraphe à l’aide du contexte de demande.

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

#### Informations supplémentaires

Vous devez appeler l’objet load() une fois les objets suivis ajoutés.

### sync()
Envoie les demandes en file d’attente à Word et renvoie un objet de promesse, qui peut être utilisé pour ajouter d’autres actions en chaîne.

#### Syntaxe
```js
requestContextObject.sync();
```

#### Paramètres
Aucun

#### Retourne
Objet de promesse.

#### Exemples

L’exemple suivant illustre la méthode de synchronisation, qui est utilisée deux fois : 1) pour charger la collection de contrôles de contenu avec la propriété de texte associée à chaque contrôle ; 2) pour désactiver le contenu du premier contrôle de contenu de la collection.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the content controls collection.
    var contentControls = context.document.contentControls;

    // Queue a command to load the content controls collection.
    contentControls.load('text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {

        if (contentControls.items.length === 0) {
            console.log("There isn't a content control in this document.");
        } else {

            // Queue a command to clear the contents of the first content control.
            contentControls.items[0].clear();
            // Synchronize the document state by executing the queued commands,
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Content control cleared of contents.');
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

## Informations de prise en charge
Utilisez l’[ensemble de conditions requises](../office-add-in-requirement-sets.md) dans les vérifications à l’exécution pour vous assurer que votre application est prise en charge par la version d’hôte de Word. Pour plus d’informations sur la configuration requise pour le serveur et l’application d’hôte Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).