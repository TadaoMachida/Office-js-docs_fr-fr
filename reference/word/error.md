# Objet OfficeExtension.Error (API JavaScript pour Word)

Représente des erreurs qui se produisent lorsque vous utilisez l’API JavaScript Word.

_S’applique à : Word 2016, Word pour iPad, Word pour Mac_

## Propriétés
| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|code|string|Obtient une valeur qui indique le type d’erreur. La valeur peut être « AccessDenied », « GeneralException », « ActivityLimitReached », « InvalidArgument », « ItemNotFound » ou « NotImplemented ». <!-- Values come from OfficeExtension.Error and Word.ErrorCodes. -->|
|debugInfo|string|Obtient une valeur qui indique ce qui s’est passé lorsque l’erreur est survenue. Cette valeur est uniquement destinée au développement/débogage.  |
|message |string| Obtient une chaîne localisée explicite qui correspond au code d’erreur.|
|name |string| Obtient une valeur qui est toujours « OfficeExtension.Error ». |
|traceMessages |string[]| Obtient un tableau de valeurs qui correspondent aux messages d’instrumentation définis avec context.trace(); |

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[toString()](#tostring)|chaîne|Renvoie le code d’erreur et le message au format suivant : « {0}: {1} », code, message.|

## Détails de méthodes

### toString()
Renvoie le code d’erreur et le message au format suivant : « {0}: {1} », code, message.

#### Syntaxe
```js
error.toString()
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

    // Queue a commmand to insert text in to the beginning of the body.
    // This will cause an OfficeExtension.Error.
    body.insertText(0);

    // Synchronize the document state by executing the queued-up commands,
    // and return a promise to indicate task completion.
    return context.sync();
})
.catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
        console.log('Error code and message: ' + error.toString());
    }
});

```

## Exemples d’accès aux propriétés

### Instrumentation de message de trace

L’exemple suivant montre comment instrumenter un lot de commandes pour déterminer où une erreur s’est produite. Le premier lot insère les deux premiers paragraphes dans le document sans provoquer d’erreurs. Le deuxième lot insère les troisième et quatrième paragraphes, mais l’appel d’insertion du cinquième paragraphe échoue. Aucune des autres commandes du lot après celle en échec n’est exécutée, y compris la commande qui ajoute le cinquième message de trace. Dans ce cas, l’erreur s’est produite après l’insertion du quatrième paragraphe et avant l’ajout du cinquième message de trace.

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a commmand to insert the paragraph at the end of the document body.
    // Start a batch of commands.
    body.insertParagraph('1st paragraph', Word.InsertLocation.end);
    // Queue a command for instrumenting this part of the batch.
    context.trace('1st paragraph successful');

    body.insertParagraph('2nd paragraph', Word.InsertLocation.end);
    context.trace('2nd paragraph successful');

    // Synchronize the document state by executing the queued-up commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        // Queue a commmand to insert the paragraph at the end of the document body.
        // Start a new batch of commands.
        body.insertParagraph('3rd paragraph', Word.InsertLocation.end);
        context.trace('3rd paragraph successful');

        body.insertParagraph('4th paragraph', Word.InsertLocation.end);
        context.trace('4th paragraph successful');

        // This command will cause an error. The trace messages in the queue up to
        // this point will be available via Error.traceMessages.
        body.insertParagraph(0, '5th paragraph', Word.InsertLocation.end);
        // Queue a command for instrumenting this part of the batch.
        // This trace message will not be set on Error.traceMessages.
        context.trace('5th paragraph successful');
    }).then(context.sync);
})
.catch(function (error) {
    if (error instanceof OfficeExtension.Error) {
        console.log('Trace messages: ' + error.traceMessages);
    }
});

// Output: "Trace messages: 3rd paragraph successful,4th paragraph successful"

```
