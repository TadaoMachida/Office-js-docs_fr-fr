# Objet ImageOcrData (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Représente les données obtenues suite à la reconnaissance optique des caractères (OCR) d’une image

## Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|ocrLanguageId|chaîne|Représente la langue de la reconnaissance optique des caractères, avec des valeurs telles que EN-US|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-imageOcrData-ocrLanguageId)|
|ocrText|chaîne|Représente le texte obtenu par reconnaissance optique des caractères de l’image|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-imageOcrData-ocrText)|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
Aucun


## Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-imageOcrData-load)|

## Détails des méthodes


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
### Exemples d’accès aux propriétés
**ocrText et ocrLanguageId**
```js
var image = null;

OneNote.run(function(ctx){
    // Get the current outline.
    var outline = ctx.application.getActiveOutline();

    // Queue a command to load paragraphs and their types.
    outline.load("paragraphs")
    return ctx.sync().
        then(function(){
            for (var i=0; i < outline.paragraphs.items.length; i++)
            {
                var paragraph = outline.paragraphs.items[i];
                if (paragraph.type == "Image")
                {
                    image = paragraph.image;
                }
            }
            if (image != null)
            {
               image.load("ocrData");
            }
            return ctx.sync();
        })
        .then(function(){
            
            // Log ocrText and ocrLanguageId
            console.log(image.ocrData.ocrText);
            console.log(image.ocrData.ocrLanguageId);
        });
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
