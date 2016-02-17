# Objet SearchResultCollection (interface API JavaScript pour Word)

Contient une collection d’objets de [plage](range.md) obtenus à la suite d’une opération de recherche.

_S’applique à : Word 2016, Word pour iPad, Word pour Mac_

## Propriétés
| Propriété   | Type|Description
|:---------------|:--------|:----------|
|Items|[SearchResult[]](searchresult.md)|Collection d’objets searchResult. En lecture seule.|

## Relations
Aucun


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
    
    // Setup the search options.
    var options = Word.SearchOptions.newObject(context);
    options.matchWildCards = true;

    // Queue a command to search the document for any string of characters after 'to'.
    var searchResults = context.document.body.search('to*', options);

    // Queue a command to load the search results and get the font property values.
    context.load(searchResults, 'font');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('Found count: ' + searchResults.items.length);

        // Queue a set of commands to change the font for each found item.
        for (var i = 0; i < searchResults.items.length; i++) {
            searchResults.items[i].font.color = 'purple';
            searchResults.items[i].font.highlightColor = 'pink';
            searchResults.items[i].font.bold = true;
        }
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync();
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
