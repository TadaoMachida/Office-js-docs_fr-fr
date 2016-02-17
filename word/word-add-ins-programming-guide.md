# Présentation de la programmation JavaScript pour les compléments Word

_S’applique à : Word 2016, Word pour iPad, Word pour Mac_

Word 2016 présente un nouveau modèle objet pour utiliser les objets de Word. Ce modèle objet complète celui déjà fourni par Office.js pour créer des compléments pour Word. Ce modèle objet est accessible via du code JavaScript hébergé par une application web.

## manifeste

La nouvelle API JavaScript pour les compléments Word utilise le même format de fichier manifeste que le modèle de complément d’Office 2013. Le fichier manifeste indique l’emplacement d’hébergement du complément, son mode d’affichage, les autorisations qui lui sont associées et d’autres informations. Pour en savoir plus sur la façon dont vous pouvez personnaliser les fichiers manifeste des compléments, cliquez [ici](https://msdn.microsoft.com/en-us/library/office/fp161044.aspx). 

Vous disposez de plusieurs options pour publier des manifestes de complément Word. Pour en savoir plus sur la façon de publier votre complément Office sur un partage réseau, un catalogue ou sur l’Office Store, cliquez [ici](https://msdn.microsoft.com/EN-US/library/office/fp123515.aspx).

## Présentation de l’API JavaScript pour Word

L’API JavaScript pour Word est chargée par Office.js. Elle fournit un ensemble d’objets de proxy JavaScript qui sont utilisés pour mettre en file d’attente un ensemble de commandes qui interagissent avec le contenu d’un document Word. Ces commandes sont exécutées sous forme de lot. Les résultats de ces commandes sont des actions appliquées au document Word, par exemple, insérer du contenu et synchroniser les objets Word avec les objets de proxy JavaScript. 

### Exécution de votre complément

Jetons un œil à ce dont vous avez besoin pour exécuter votre complément. Tous les compléments doivent disposer d’un gestionnaire d’événements Office.initialize.  Pour plus d’informations sur l’initialisation du complément, voir [Présentation de l’API](https://msdn.microsoft.com/EN-US/library/fp160953.aspx) .  

Votre complément Word s’exécute en insérant une fonction dans la méthode Word.run(). La fonction transmise dans la méthode d’exécution doit contenir un argument de contexte. Cet [objet de contexte](word-add-ins-javascript-reference/requestcontext.md) est différent de celui que vous obtenez de l’objet Office, même s’il a la même finalité, à savoir interagir avec l’environnement d’exécution de Word. L’objet de contexte permet d’accéder au modèle objet JavaScript de Word. Examinons les commentaires et le code d’un complément Word de base :

**Exemple 1. Initialisation et exécution d’un complément Word**

```javascript
    (function () {
        "use strict";

        // The initialize event handler is run each time the page is loaded.
        Office.initialize = function (reason) {
            
            // Checks for the DOM to load using the jQuery ready function.
            $(document).ready(function () {
                // Set your initialization code. You can use the reason 
                // argument to determine how the add-in was loaded.
                // You can also load saved settings from the Office object.
            });
        };

        // Run a batch operation against the Word object model.
        // Use the context argument to get access to the Word document.
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;
        })
    })();
```

L’exemple 1 présente le code de base nécessaire pour créer un complément Word. Il initialise Office.js et contient une méthode d’exécution pour l’interaction avec le document Word.

### Objets de proxy

Le modèle objet JavaScript pour Word est associé de façon relativement libre avec les objets dans Word. Les objets JavaScript pour Word sont des objets de proxy correspondant aux objets réels d’un document Word. Toutes les actions effectuées sur les objets de proxy ne sont pas réalisées dans Word et l’état du document Word n’est pas répercuté sur les objets de proxy tant que cet état n’a pas été synchronisé. L’état de document est synchronisé lors de l’exécution de la méthode context.sync(). Celle-ci exécute l’ensemble des commandes de la file d’attente pour chaque objet de proxy.  L’exemple 2 présente la création d’un objet Body de proxy et une file de commandes permettant de charger la propriété de texte sur l’objet Body de proxy, puis la synchronisation du corps dans le document Word avec l’objet de proxy correspondant. 

**Exemple 2. Synchronisation du corps du document avec l’objet de proxy correspondant.**

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        // The body object hasn't been set with any property values. 
        var body = context.document.body;

        // Queue a command to load the text property for the proxy document body object.
        context.load(body, 'text');

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });  
    })
```

### File d’attente de commandes

Les objets de proxy Word utilisent des méthodes pour accéder au modèle objet et le mettre à jour. Ces méthodes sont exécutées l’une après l’autre, dans l’ordre dans lequel elles ont incluses dans la file d’attente du lot. Un lot de commandes est constitué avant l’appel de la méthode context.sync(). Toutes les commandes en attente dans tous les objets qui utilisent le contexte sont exécutées.  

L’exemple 3 montre comment fonctionne la file d’attente de commandes. Lorsque la méthode context.sync() est appelée, la [commande visant à charger](Word%20Add-ins%20JavaScript%20Reference/loadoption.md) le corps du texte est tout d’abord exécutée dans Word. C’est ensuite la commande visant à insérer du texte dans le corps de Word qui est appliquée. Les résultats sont alors renvoyés vers l’objet Body de proxy. La valeur de la propriété body.text dans le code JavaScript Word est la même que celle du corps du document de Word <u>avant</u> l’insertion du texte dans le document Word. 

**Exemple 3. Exécution d’un lot de commandes.**

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to load the text in the proxy body object.
        context.load(body, 'text');

        // Queue a command to insert text into the end of the Word document body.
        body.insertText('This is text inserted after loading the body.text property',
                        Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });  
    })
```

## Donnez-nous votre avis.

Votre avis compte beaucoup pour nous. 

* Consultez les documents et signalez-nous toute question ou tout problème à leur propos en [soumettant une question](https://github.com/OfficeDev/office-js-docs/issues) directement dans ce référentiel.
* Faites-nous part de vos expériences de programmation, de ce que vous souhaiteriez voir dans les futures versions, de vos questions sur les exemples de code, etc. Passez par [ce site](http://officespdev.uservoice.com/) pour soumettre vos suggestions et vos idées.


## Ressources supplémentaires

* [Compléments Word](word-add-ins.md)
* [Référence des API JavaScript pour les compléments Word](word-add-ins-javascript-reference.md)
* [Compléments Office](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Commencer à utiliser les compléments Office](http://dev.office.com/getting-started/addins)
* &lt;a herf="https://github.com/OfficeDev?utf8=%E2%9C%93&amp;query=Word"&gt;Compléments Word sur GitHub&lt;/a&gt;
* [Explorateur d’extraits de code pour Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)

