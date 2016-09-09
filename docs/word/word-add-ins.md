# Créer votre premier complément Word

_S’applique à : Word 2016, Word pour iPad, Word pour Mac_

L’interface API JavaScript pour Word est comprise dans le modèle de programmation des compléments Office, qui vise à étendre les applications Office. Ce modèle utilise des applications web pour héberger votre extension dans Word. Vous pouvez désormais prolonger les fonctionnalités de Word avec la plateforme web ou la langue de votre choix.

Un complément Word est exécuté à l’intérieur de Word et peut interagir avec le contenu du document à l’aide des interfaces API JavaScript pour Word disponibles dans Word 2016. Ce système utilise deux composants pour créer un complément : 1) une application web que vous pouvez héberger n’importe où, et 2) le [fichier manifeste de complément](../../docs/overview/add-in-manifests.md) que Word utilise pour repérer l’emplacement où votre application web est hébergée (pour en savoir plus sur les autres fonctions du fichier manifeste, consultez la [vue d’ensemble de la programmation](word-add-ins-programming-overview.md)).

>**Complément Word = manifest.xml + application web**

### Configuration
Dans cette section, vous allez créer une simple application web, ainsi que le fichier manifeste correspondant. L’application web vous permettra d’ajouter du texte réutilisable dans le document Word.

1- Créez un dossier nommé BoilerplateAddin sur votre disque local (par exemple C:\\BoilerplateAddin). Vous devrez enregistrer tous les fichiers créés au cours des étapes qui suivent dans ce dossier.

2- Créez un fichier nommé home.html pour l’affichage du complément. Le complément comportera trois boutons qui, lorsqu’ils seront sélectionnés, ajouteront du texte réutilisable. Collez le code suivant dans le fichier home.html.

```html
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Boilerplate text app</title>
        <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
        <script src="home.js" type="text/javascript"></script>
        </head>
        <body>
            <div>
                    <h1>Welcome</h1>
            </div>
            <div>
                    <p>This sample shows how to add boilerplate text to a document by using the Word JavaScript API.</p>
                    <br />
                    <h3>Try it out</h3>
                    <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                    <button id="checkhov">Add quote from Anton Chekhov</button>
                    <button id="proverb">Add Chinese proverb</button>
            </div>
            <h3><div id="supportedVersion"/></h3>
        </body>
    </html>
```

3- Créez un fichier nommé home.js et collez-y le code suivant. Il contient le code d’initialisation et l’ensemble du code nécessaire au complément pour apporter des modifications au document Word. Ce code insère du texte en fonction de la position du curseur ou des éléments sélectionnés dans le document Word.

```javascript
    (function () {
        "use strict";

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
```

4- Créez un fichier XML nommé BoilerplateManifest.xml et collez-y le code. Il s’agit du fichier manifeste que Word utilise pour repérer des informations sur un complément, telles que son emplacement ou son nom complet.
```xml
<?xml version="1.0" encoding="UTF-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xsi:type="TaskPaneApp">
        <Id>2b88100c-656e-4bab-9f1e-f6731d86e464</Id>
        <Version>1.0.0.0</Version>
        <ProviderName>Microsoft</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
        <DisplayName DefaultValue="Boilerplate content" />
        <Description DefaultValue="Insert boilerplate content into a Word document." />
        <Hosts>
            <Host Name="Document"/>
        </Hosts>
        <DefaultSettings>
            <SourceLocation DefaultValue="\\MyShare\boilerplate\home.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
```

5- Générez un GUID et utilisez-le pour remplacer la valeur de l’élément <code>OfficeApp/Id</code>.

6- Enregistrez tous les fichiers. Vous venez d’écrire votre premier complément Word.

7- Copiez les fichiers home.js, home.html et BoilerplateManifest.xml vers un [dossier partagé sur le réseau](https://technet.microsoft.com/en-us/library/cc770880.aspx) (Windows) ou hébergez-les sur un serveur local (Mac).

8- Modifiez l’élément [SourceLocation](../../reference/manifest/sourcelocation.md) dans le fichier BoilerplateManifest.xml afin qu’il pointe vers l’emplacement du fichier home.html.

À ce stade, votre premier complément est déployé. Vous devez maintenant indiquer à Word où trouver le complément.

#### Faire un essai dans Word 2016 pour Windows

1. Lancez Word et ouvrez un document.
2. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
3. Choisissez l’onglet **Fichier**, puis choisissez **Options**.
4. Choisissez **Catalogues de compléments approuvés**.
5. Dans la zone **URL du catalogue**, entrez le chemin d’accès au partage de dossier contenant le fichier BoilerplateManifest.xml, puis choisissez **Ajout d’un catalogue**.
6. Activez la case à cocher **Afficher dans le menu**, puis cliquez sur **OK**.
7. Un message vous informe que vos paramètres seront appliqués lors du prochain démarrage d’Office. Fermez et redémarrez Word.

Vous pouvez à présent exécuter le complément que vous avez créé. Pour le voir à l’œuvre, procédez comme suit :

1. Ouvrez un document Word.
2. Dans l’onglet **Insertion** de Word 2016, choisissez **Mes compléments**.
3. Sélectionnez l’onglet **Dossier partagé**.
4. Choisissez **Contenu réutilisable**, puis sélectionnez **Insérer**.
5. Le complément est chargé dans un volet de tâches. Reportez-vous à la figure 1 pour voir l’aspect du complément une fois chargé.
6. Sélectionnez les boutons pour entrer du texte réutilisable dans le document Word.


### Faire un essai dans Word 2016 pour Mac

Vous pouvez à présent exécuter le complément que vous avez créé. Pour le voir à l’œuvre, procédez comme suit :

1. Créez un dossier nommé « wef » dans Users/Library/Containers/com.microsoft.word/Data/Documents/
2. Placez le fichier manifeste, BoilerplateManifest.xml, dans le dossier « wef » (Users/Library/Containers/com.microsoft.word/Data/Documents/wef)
3. Ouvrez Word 2016 sur le Mac et cliquez sur l’onglet Insertion, puis sur la liste déroulante Mes compléments. Vous devez voir le complément dans la liste déroulante. Sélectionnez-le pour le charger.

__Figure 1. Complément de contenu réutilisable chargé dans Word__
![Image de l’application Word une fois le complément réutilisable chargé.](../../images/boilerplateAddin.png "Un simple complément Word permettant d’entrer du texte réutilisable.")

## Donnez-nous votre avis.

Votre avis compte beaucoup pour nous.

* Consultez les documents et signalez-nous toute question ou tout problème à leur propos en [soumettant une question](https://github.com/OfficeDev/office-js-docs/issues).
* Faites-nous part de vos expériences de programmation et de ce que vous souhaiteriez voir dans les futures versions ou les exemples de code. Passez par [le site UserVoice](http://officespdev.uservoice.com/) pour soumettre vos suggestions et vos idées.

## Ressources supplémentaires

* [Commencer à utiliser les compléments Office](https://dev.office.com/getting-started/addins?product=word)
* [Compléments Word sur GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Word)
