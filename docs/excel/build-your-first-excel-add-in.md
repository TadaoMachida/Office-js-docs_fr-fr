# Création de votre premier complément Excel

Cet article décrit comment utiliser l’API JavaScript Excel pour créer un complément pour Excel 2016 ou Excel Online. La procédure suivante vous guide tout au long du processus de création d’un simple complément de volet Office permettant de charger des données dans une feuille de calcul et de créer un graphique de base dans Excel 2016.

![Complément de rapport sur les ventes trimestrielles](../../images/QuarterlySalesReport_report.PNG)


Vous devez commencer par créer une application web en utilisant HTML et JQuery. Ensuite, vous devez créer un fichier manifeste XML qui indique l’endroit où vous souhaitez localiser votre application web et la façon dont elle doit apparaître dans Excel.


### Le coder

1- Créez un dossier sur votre lecteur local nommé RapportVentesTrimestrielles (par exemple, C:\\RapportVentesTrimestrielles). Vous devrez enregistrer tous les fichiers créés au cours des étapes qui suivent dans ce dossier.

2- Créez la page HTML qui sera chargée dans le complément de volet de tâches. Nommez le fichier **Home.html** et collez le code ci-dessous dans le fichier.

```html

    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Quarterly Sales Report</title>

        <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>

        <link href="Office.css" rel="stylesheet" type="text/css" />

        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>

        <link href="Common.css" rel="stylesheet" type="text/css" />
        <script src="Notification.js" type="text/javascript"></script>

        <script src="Home.js" type="text/javascript"></script>

        <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
        <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">

    </head>
    <body class="ms-font-m">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>This sample shows how to load some sample data into the worksheet, and then create a chart using the Excel JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button class="ms-Button" id="load-data-and-create-chart">Click me!</button>
            </div>
        </div>
    </body>
    </html>

```

3- Créez un fichier nommé **Common.css** pour stocker vos styles personnalisés et collez le code ci-dessous dans le fichier.

```css
    /* Common app styling */

    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; /* Fixed header height */
        overflow: hidden; /* Disable scrollbars for header */
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px; /* Same value as #content-header's height */
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; /* Enable scrollbars within main content section */
    }

    .padding {
        padding: 15px;
    }

    #notification-message {
        background-color: #818285;
        color: #fff;
        position: absolute;
        width: 100%;
        min-height: 80px;
        right: 0;
        z-index: 100;
        bottom: 0;
        display: none; /* Hidden until invoked */
    }

        #notification-message #notification-message-header {
            font-size: medium;
            margin-bottom: 10px;
        }

        #notification-message #notification-message-close {
            background-image: url("../../images/Close.png");
            background-repeat: no-repeat;
            width: 24px;
            height: 24px;
            position: absolute;
            right: 5px;
            top: 5px;
            cursor: pointer;
        }


```

4- Créez un fichier qui contiendra la logique de programmation pour le complément dans jQuery. Nommez le fichier **Home.js** et collez le script suivant dans le fichier.

```js

    (function () {
        "use strict";

        // The initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
                app.initialize();

                $('#load-data-and-create-chart').click(loadDataAndCreateChart);
            });
        };

        // Load some sample data into the worksheet and then create a chart
        function loadDataAndCreateChart() {
            // Run a batch operation against the Excel object model
            Excel.run(function (ctx) {

                // Create a proxy object for the active worksheet
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();

                //Queue commands to set the report title in the worksheet
                sheet.getRange("A1").values = "Quarterly Sales Report";
                sheet.getRange("A1").format.font.name = "Century";
                sheet.getRange("A1").format.font.size = 26;

                //Create an array containing sample data
                var values = [["Product", "Qtr1", "Qtr2", "Qtr3", "Qtr4"],
                              ["Frames", 5000, 7000, 6544, 4377],
                              ["Saddles", 400, 323, 276, 651],
                              ["Brake levers", 12000, 8766, 8456, 9812],
                              ["Chains", 1550, 1088, 692, 853],
                              ["Mirrors", 225, 600, 923, 544],
                              ["Spokes", 6005, 7634, 4589, 8765]];

                //Queue a command to write the sample data to the specified range
                //in the worksheet and bold the header row
                var range = sheet.getRange("A2:E8");
                range.values = values;
                sheet.getRange("A2:E2").format.font.bold = true;

                //Queue a command to add a new chart
                var chart = sheet.charts.add("ColumnClustered", range, "auto");

                //Queue commands to set the properties and format the chart
                chart.setPosition("G1", "L10");
                chart.title.text = "Quarterly sales chart";
                chart.legend.position = "right"
                chart.legend.format.fill.setSolidColor("white");
                chart.dataLabels.format.font.size = 15;
                chart.dataLabels.format.font.color = "black";
                var points = chart.series.getItemAt(0).points;
                points.getItemAt(0).format.fill.setSolidColor("pink");
                points.getItemAt(1).format.fill.setSolidColor('indigo');

                //Run the queued commands, and return a promise to indicate task completion
                return ctx.sync();
            })
              .then(function () {
                  app.showNotification("Success");
                  console.log("Success!");
              })
            .catch(function (error) {
                // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
                app.showNotification("Error: " + error);
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
```


5- Créez un fichier qui contiendra la logique de programmation pour fournir des notifications dans le complément en cas d’erreur. Ces notifications sont utiles pour les tâches de débogage. Nommez le fichier **Notification.js** et collez le script suivant dans le fichier.

```js

    /* Notification functionality */

    var app = (function () {
        "use strict";

        var app = {};

        // Initialization function (to be called from each page that needs notification)
        app.initialize = function () {
            $('body').append(
                '<div id="notification-message">' +
                    '<div class="padding">' +
                        '<div id="notification-message-close"></div>' +
                        '<div id="notification-message-header"></div>' +
                        '<div id="notification-message-body"></div>' +
                    '</div>' +
                '</div>');

            $('#notification-message-close').click(function () {
                $('#notification-message').hide();
            });


            // After initialization, expose a common notification function
            app.showNotification = function (header, text) {
                $('#notification-message-header').text(header);
                $('#notification-message-body').text(text);
                $('#notification-message').slideDown('fast');
            };
        };

        return app;
    })();
```

6- Créez un fichier manifeste XML pour indiquer l’emplacement de votre application web et la façon dont vous voulez qu’elle apparaisse dans Excel. Nommez le fichier **RapportVentesTrimestrielles_Manifeste.xml** et collez le code XML suivant dans le fichier.

```xml
    <?xml version="1.0" encoding="UTF-8"?>
    <!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
      <Id>ab2991e7-fe64-465b-a2f1-c865247ef434</Id>
      <Version>1.0.0.0</Version>
      <ProviderName>Microsoft</ProviderName>
      <DefaultLocale>en-US</DefaultLocale>
      <DisplayName DefaultValue="Quarterly Sales Report Sample" />
      <Description DefaultValue="Quarterly Sales Report Sample"/>
      <Capabilities>
        <Capability Name="Workbook" />
      </Capabilities>
      <DefaultSettings>
        <SourceLocation DefaultValue="\\MyShare\QuarterlySalesReport\Home.html" />
      </DefaultSettings>
      <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
```

7- Générez un GUID à l’aide d’un générateur en ligne de votre choix. Ensuite, remplacez la valeur de la balise **Id** indiquée à l’étape précédente par ce GUID.

8-  Enregistrez tous les fichiers. Vous venez d’écrire votre premier complément Excel.

### Essayez !

Pour déployer et tester votre complément, le plus simple consiste à copier les fichiers sur un partage réseau.

1- Créez un dossier sur un partage réseau (par exemple, \\\MyShare\\RapportVentesTrimestrielles) et copiez tous les fichiers dans ce dossier.

2- Modifiez l’élément **SourceLocation** du fichier manifeste afin qu’il pointe vers l’emplacement de partage de la page .html de l’étape 1.

3- Copiez le fichier manifeste (RapportVentesTrimestrielles_Manifeste.xml) sur un partage réseau (par exemple, \\\MyShare\\MesManifestes).

4- Maintenant, ajoutez l’emplacement de partage qui contient le fichier manifeste sous forme de catalogue d’applications approuvées dans Excel. Lancez Excel et ouvrez une feuille de calcul vide.

5-  Choisissez l’onglet **Fichier**, puis choisissez **Options**.

6-  Choisissez **Centre de gestion de la confidentialité**, puis cliquez sur le bouton **Paramètres du Centre de gestion de la confidentialité**.

7-  Choisissez **Catalogues de compléments approuvés**.

8-  Dans la zone **URL du catalogue**, entrez le chemin d’accès au partage réseau que vous avez créé à l’étape 3, puis choisissez **Ajouter un catalogue**. Cochez la case **Afficher dans le menu**, puis cliquez sur **OK**. Un message vous informe que vos paramètres seront appliqués lors du prochain démarrage d’Office.

9-  Maintenant, testez et exécutez le complément. Dans l’onglet **Insertion** d’Excel 2016, choisissez **Mes compléments**.

10-  Dans la boîte de dialogue **Compléments Office**, choisissez **Dossier partagé**.

11-  Choisissez **Exemple de rapport sur les ventes trimestrielles**>**Insertion**. Le complément s’ouvre dans un volet Office à droite de la feuille de calcul active, comme indiqué dans l’illustration suivante.

 ![Complément de rapport sur les ventes trimestrielles](../../images/QuarterlySalesReport_taskpane.PNG)

12-  Cliquez sur le bouton **Cliquez ici !** pour afficher les données et le graphique à l’intérieur de la feuille de calcul, comme indiqué dans l’illustration suivante.  Pour mettre à jour le graphique dynamiquement, il vous suffit de modifier les données de la plage.

![Complément de rapport sur les ventes trimestrielles](../../images/QuarterlySalesReport_report.PNG)


### Ressources supplémentaires

*  [Présentation de la programmation JavaScript pour les compléments Excel](excel-add-ins-javascript-programming-overview.md)
*  [Explorateur d’extraits de code pour Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
*  [Exemples de code pour les compléments Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
*  [Référence de l’API JavaScript pour les compléments Excel](excel-add-ins-javascript-api-reference.md)
