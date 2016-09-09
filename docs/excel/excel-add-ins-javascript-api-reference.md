# Référence de l’API JavaScript d’Excel

Vous pouvez utiliser l’API JavaScript d’Excel pour créer des compléments pour Excel 2016. La liste suivante affiche les objets de haut niveau Excel qui sont disponibles dans l’API. Chaque page d’objet contient une description des propriétés, des relations et des méthodes disponibles sur l’objet. Cliquez sur les liens suivants pour en savoir plus.

* [Workbook](../../reference/excel/workbook.md) : objet de niveau supérieur qui contient les objets de classeur associés tels que les feuilles de calcul, les tableaux, les plages, etc. Il permet également d’établir la liste des références associées.
* [Worksheet](../../reference/excel/worksheet.md) : membre de la collection de feuilles de calcul. La collection de feuilles de calcul contient tous les objets de feuille de calcul d’un classeur.
    * [WorksheetCollection](../../reference/excel/worksheetcollection.md) : collection de tous les objets de classeur qui font partie du classeur.
* [Range](../../reference/excel/range.md) : représente une cellule, une ligne, une colonne ou une sélection de cellules contenant un ou plusieurs blocs contigus de cellules.
* [Table](../../reference/excel/table.md) : représente une collection de cellules organisées conçue pour faciliter la gestion des données.
    * [TableCollection](../../reference/excel/tablecollection.md) : collection de tableaux d’un classeur ou d’une feuille de calcul.
    * [TableColumnCollection](../../reference/excel/tablecolumncollection.md) : collection de toutes les colonnes d’un tableau.
    * [TableRowCollection](../../reference/excel/tablerowcollection.md) : collection de toutes les lignes d’un tableau.
* [Chart](../../reference/excel/chart.md) : représente un objet de graphique d’une feuille de calcul, qui est une représentation visuelle de données sous-jacentes.
    * [ChartCollection](../../reference/excel/chartcollection.md) : collection de graphiques d’une feuille de calcul.
* [TableSort](../../reference/excel/tablesort.md) : représente un objet qui tri les opérations sur les objets Table.
* [RangeSort](../../reference/excel/rangesort.md) : représente un objet qui tri les opérations sur les objets Range.
* [Filter](../../reference/excel/filter.md) : représente un objet filter qui gère le filtrage de colonne d’un tableau.
* [WorksheetProtection](../../reference/excel/worksheetprotection.md) : représente la protection d’un objet de la feuille.
* [WorksheetFunction](../../reference/excel/functions.md) : représente un conteneur pour les fonctions de feuille de calcul Microsoft Excel que vous pouvez appeler via JavaScript.
* [NamedItem](../../reference/excel/nameditem.md) : représente un nom défini pour une plage de cellules ou une valeur. Ces noms peuvent comprendre des objets primitifs nommés, des objets de plage, etc.
    * [NamedItemCollection](../../reference/excel/nameditemcollection.md) : collection d’objets NamedItem d’un classeur.
* [Binding](../../reference/excel/binding.md) : classe abstraite qui représente une liaison à une section du classeur.
    * [Binding Collection](../../reference/excel/bindingcollection.md) : collection de tous les objets de liaison qui font partie du classeur.
* [TrackedObjectCollection](../../reference/excel/trackedobjectscollection.md) : permet aux compléments de gérer une référence d’objet de plage sur plusieurs lots sync().
* [Request Context](../../reference/excel/requestcontext.md) : l’objet de contexte de demande facilite les demandes auprès de l’application Excel.


##### Ressources supplémentaires

*  [Présentation de la programmation JavaScript pour les compléments Excel](excel-add-ins-javascript-programming-overview.md)
*  [Création de votre premier complément Excel](build-your-first-excel-add-in.md)
*  [Explorateur d’extraits de code pour Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)

