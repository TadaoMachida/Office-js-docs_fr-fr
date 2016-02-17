# Référence de l’API JavaScript pour les compléments Excel

_S’applique à : Excel 2016, Office 2016_

Les liens ci-dessous renvoient aux objets Excel de niveau supérieur disponibles dans les API. Chaque page d’objet contient une description des propriétés, des relations et des méthodes disponibles sur l’objet. Explorez les liens ci-dessous pour en savoir plus.
	
* [Workbook](resources/workbook.md) : objet de niveau supérieur qui contient les objets de classeur associés tels que les feuilles de calcul, les tableaux, les plages, etc. Il permet également d’établir la liste des références associées. 
* [Worksheet](resources/worksheet.md) : membre de la collection de feuilles de calcul. La collection de feuilles de calcul contient tous les objets de feuille de calcul d’un classeur.
	* [WorksheetCollection](resources/worksheetcollection.md) : collection de tous les objets de classeur qui font partie du classeur. 
* [Range](resources/range.md) : représente une cellule, une ligne, une colonne ou une sélection de cellules contenant un ou plusieurs blocs contigus de cellules.  
* [Table](resources/table.md) : représente une collection de cellules organisées conçue pour faciliter la gestion des données. 
	* [TableCollection](resources/tablecollection.md) : collection de tableaux d’un classeur ou d’une feuille de calcul. 
	* [TableColumnCollection](resources/tablecolumncollection.md) : collection de toutes les colonnes d’un tableau. 
	* [TableRowCollection](resources/tablerowcollection.md) : collection de toutes les lignes d’un tableau. 
* [Chart](resources/chart.md) : représente un objet de graphique d’une feuille de calcul, qui est une représentation visuelle de données sous-jacentes.   
	* [ChartCollection](resources/chartcollection.md) : collection de graphiques d’une feuille de calcul.	
* [NamedItem](resources/nameditem.md) : représente un nom défini pour une plage de cellules ou une valeur. Ces noms peuvent comprendre des objets primitifs nommés, des objets de plage, etc.
	* [NamedItemCollection](resources/nameditemcollection.md) : collection d’objets NamedItem d’un classeur.
* [Binding](resources/binding.md) : classe abstraite qui représente une liaison à une section du classeur.
	* [Binding Collection](resources/bindingcollection.md) : collection de tous les objets de liaison qui font partie du classeur. 
* [TrackedObjectCollection](resources/trackedobjectscollection.md) : permet aux compléments de gérer une référence d’objet de plage sur plusieurs lots sync(). 
* [Request Context](resources/requestcontext.md) : l’objet de contexte de demande facilite les demandes auprès de l’application Excel.


##### Ressources supplémentaires

*  [Présentation de la programmation JavaScript pour les compléments Excel](excel-add-ins-programming-overview.md)
*  [Création de votre premier complément Excel](build-your-first-excel-add-in.md)
*  [Explorateur d’extraits de code pour Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
*  [Exemples de code pour les compléments Excel](excel-add-ins-code-samples.md) 


