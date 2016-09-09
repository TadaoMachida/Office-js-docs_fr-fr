# Options de chargement d’objet 

Représente un objet qui peut être transmis à la méthode de chargement pour spécifier l’ensemble de propriétés et de relations à charger lors de l’exécution de la méthode sync() qui synchronise les états entre les objets OneNote et les objets de proxy JavaScript correspondants dans le complément. Cet objet utilise des options telles que les paramètres des propriétés select et expand afin de spécifier l’ensemble de propriétés à charger sur l’objet et autorise la pagination sur la collection.

Vous pouvez également fournir une chaîne ou un tableau qui contient les propriétés et les relations à charger, tel qu’illustré dans l’exemple suivant.

```js   
object.load('<var1>,<relationship1/var2>');

// Pass the parameter as an array.
object.load(["var1", "relationship1/var2"]);
```

## Propriétés
| Propriété     | Type   |Description|
|:---------------|:--------|:----------|
|select|object|Fournissez un tableau ou une liste de noms de paramètres/relations (en les séparant par des virgules) à charger lors d’un appel synchronisé. Par exemple, "propriété1, relation1", ["propriété1", "relation1"]. Facultatif.|
|expand|object|Fournissez un tableau ou une liste de noms de relations (en les séparant par des virgules) à charger lors d’un appel synchronisé. Par exemple, "relation1, relation2", [ "relation1", "relation2"]. Facultatif.|
|top|int|Indiquez le nombre d’éléments de la collection demandée à inclure dans le résultat. Facultatif.|
|skip|int|Indiquez le nombre d’éléments de la collection devant être ignorés et exclus du résultat. Si une valeur est définie pour `top`, la sélection du résultat démarre une fois que le nombre spécifié d’éléments a été ignoré. Facultatif.|

#### Exemples

L’exemple permet d’obtenir le titre de page et le niveau de retrait des cinq premières pages dans la section active.

```js
OneNote.run(function (context) { 
    
    // Get the pages in the current section.
    var pages = context.application.getActiveSection().pages;
            
    // Queue a command to load the pages.           
    pages.load({ "select":"title,pageLevel", "top":5, "skip":0 });
    return context.sync()
        .then(function() {
            
            // Iterate through the collection of pages.    
            $.each(pages.items, function(index, page) {
                
                // Show some properties.
                console.log("Page title: " + page.title);
                console.log("Indentation level: " + page.pageLevel);
                
            });
        }).catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        })
    });
```
