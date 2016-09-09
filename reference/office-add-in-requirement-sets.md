
# Ensemble de conditions requises pour les compléments Office

Les ensembles de conditions requises sont des groupes nommés de membres d’API. Les compléments Office utilisent des ensembles de conditions spécifiés dans le manifeste ou utilisant une vérification à l’exécution pour déterminer si un hôte Office prend en charge les API nécessaires au complément. Pour plus d’informations, voir l’article sur la [spécification des conditions requises pour les API et les hôtes Office](../docs/overview/specify-office-hosts-and-api-requirements.md).

Pour obtenir une vue générale de la prise en charge des compléments par l’hôte Office, voir la page sur la [disponibilité des compléments Office sur les plateformes et les hôtes](https://dev.office.com/add-in-availability).

## Ensembles de conditions requises


Le tableau suivant répertorie les noms des ensembles de conditions requises, les méthodes de chaque ensemble et les applications hôtes d’Office qui les prennent en charge, ainsi que le numéro de version de l’API.

Pour plus d’informations sur les ensembles de conditions requises pour Outlook, voir la page de [présentation des ensembles de conditions requises pour l’API Outlook](./outlook/tutorial-api-requirement-sets.md).

|  Nom de l’ensemble  |  Version  |  Hôte Office  |  Méthodes dans l’ensemble  |
|:-----|-----|:-----|:-----|
| ExcelApi   | 1.2 | Excel 2016<br>Excel Online<br>Excel pour iPad<br>|Protection de la feuille de calcul<br>Objet Worksheet Functions<br>Tri<br>Filtrer<br>Style de référence R1C1<br>Fusionner des cellules<br>Ajuster la hauteur de ligne et la largeur de colonne<br>Chart.getImage()<br>Range.getUsedRange(valuesOnly)|
| ExcelApi   | 1.1 | Excel 2016<br>Excel Online<br>Excel pour iPad<br>|Tous les éléments dans l’espace de noms Excel|
| WordApi    | 1.2 | Word 2016<br>Word 2016 pour Mac<br>Word pour iPad<br>Word Online (aperçu) | Tous les éléments dans l’espace de noms Word. Les méthodes suivantes ont été ajoutées à cette version de WordApi :<br>Body.select(selectionMode)<br>Body.insertInlinePictureFromBase64(base64EncodedImage, insertLocation)<br>contentControl.select(selectionMode)<br>contentControl.insertInlinePictureFromBase64(base64EncodedImage, insertLocation)<br>inlinePicture.paragraph<br>inlinePicture.delete<br>inlinePicture.insertBreak(breakType, insertLocation)<br>inlinePicture.insertFileFromBase64(base64file, insertLocation)<br>inlinePicture.insertHtml(html, insertLocation)<br>inlinePicture.insertInlinePictureFromBase64(base64file, insertLocation)<br>inlinePicture.insertOoxml(ooxml, insertLocation)<br>inlinePicture.insertParagraph(paragraphText, insertLocation)<br>inlinePicture.insertText(text, insertLocation)<br>inlinePicture.select(selectionMode)<br>paragraph.select(selectionMode)<br>range.inlinePictures<br>range.select(selectionMode)<br>range.insertInlinePictureFomBase64(base64EcodedImage, insertLocation)|
| WordApi    | 1.1 | Word 2016<br>Word 2016 pour Mac<br>Word pour iPad<br>|Tous les éléments de l’espace de noms Word à l’exception des membres d’API qui ont été ajoutés à la WordApi 1.2 et versions ultérieures, lesquels sont répertoriés ci-dessus.|
| ActiveView | 1.1 | PowerPoint<br>PowerPoint Online|Document.getActiveViewAsync|
| BindingEvents  | 1.1 | Applications web Access<br>Excel<br>Excel Online<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Binding.addHanderAsync<br>Binding.removeHanderAsync|
| CompressedFile    | 1.1 |PowerPoint<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad<br/>Excel Online<br/>PowerPoint Online|Prend en charge la sortie au format Office Open XML (OOXML) sous la forme d’un tableau d’octets<br>(Office.FileType.Compressed) lorsque vous utilisez la méthode Document.getFileAsync.|
| CustomXmlParts    | 1.1 |Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|CustomXmlNode.getNodesAsync<br>CustomXmlNode.getNodeValueAsync<br>CustomXmlNode.getXmlAsync<br>CustomXmlNode.setNodeValueAsync<br>CustomXmlNode.setXmlAsync<br>CustomXmlPart.addHandlerAsync<br>CustomXmlPart.deleteAsync<br>CustomXmlPart.getNodesAsync<br>CustomXmlPart.getXmlAsync<br>CustomXmlPart.removeHandlerAsync<br>CustomXmlParts.addAsync<br>CustomXmlParts.getByIdAsync<br>CustomXmlParts.getByNamespaceAsync<br>CustomXmlPrefixMappings.addNamespaceAsync<br>CustomXmlPrefixMappings.getNamespaceAsync<br>CustomXmlPrefixMappings.getPrefixAsync|
| DialogAPI | 1.1 | Excel<br>PowerPoint<br>Word 2016<br>Outlook|Office.context.ui.displayDialogAsync()<br>Office.context.ui.messageParent()<br>Office.context.ui.close()|
| DocumentEvents    | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Document.addHandlerAsync<br>Document.removeHandlerAsync|
| Fichier  | 1.1 | PowerPoint<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad<br>PowerPoint Online|Document.getFileAsync<br>File.closeAsync<br>File.getSliceAsync|
| HtmlCoercion  | 1.1 | Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Prend en charge le forçage au format HTML (Office.CoercionType.Html) lors de la lecture et de l’écriture des données à l’aide des méthodes Document.getSelectedDataAsync,<br>Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| ImageCoercion | 1.1 | Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Prise en charge de la conversion en une image (Office.CoercionType.Image) lors de l’écriture des données à l’aide de la méthode Document.setSelectedDataAsync.|
| Boîte aux lettres   |   | Outlook pour Windows<br>Outlook pour le web<br>Outlook pour Mac<br>Outlook Web App |Voir [Présentation de l’ensemble de conditions requises pour les API Outlook](./outlook/tutorial-api-requirement-sets.md)|
| MatrixBindings    | 1.1 | Excel<br>Excel Online<br>Word|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncMatrix<br>Binding.getDataAsyncMatrix<br>Binding.setDataAsync|
| MatrixCoercion    | 1.1 | Excel<br>Excel Online<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Prise en charge du forçage de type sur la structure de données (Office.CoercionType.Matrix) « matrice » (tableau de tableaux) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| OoxmlCoercion | 1.1 | Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Prise en charge du forçage de type au format Open Office XML (OOXML) (Office.CoercionType.Ooxml) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| PartialTableBindings  | 1.1 | Applications web Access||
| PdfFile   | 1.1 | PowerPoint<br/>PowerPoint Online<br/>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Prend en charge la sortie au format PDF (Office.FileType.Pdf)<br>lorsque vous utilisez la méthode Document.getFileAsync.|
| Sélection | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Document.getSelectedDataAsync<br>Document.setSelectedDataAsync|
| Paramètres  | 1.1 | Applications web Access<br>Excel<br>Excel Online<br>PowerPoint<br>PowerPoint Online<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Settings.get<br>Settings.remove<br>Settings.saveAsync<br>Settings.set|
| TableBindings | 1.1 | Applications web Access<br>Excel<br>Excel Online<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncTable<br>Binding.addColumnsAsyncTable<br>Binding.addRowsAsyncTable<br>Binding.deleteAllDataValuesAsyncTable<br>Binding.getDataAsyncTable<br>Binding.setDataAsync|
| TableCoercion | 1.1 | Applications web Access<br>Excel<br>Excel Online<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Prise en charge du forçage de type sur la structure de données « tableau » (Office.CoercionType.Table) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| TextBindings  | 1.1 | Excel<br>Excel Online<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Bindings.addFromNamedItemAsync<br>Bindings.addFromSelectionAsync<br>Bindings.getAllAsync<br>Bindings.getByIdAsync<br>Bindings.releaseByIdAsyncText<br>Binding.getDataAsyncText<br>Binding.setDataAsync|
| TextCoercion  | 1.1 | Excel<br>Excel Online<br>PowerPoint<br>Project<br>Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad|Prise en charge du forçage de type au format texte (Office.CoercionType.Text) lors de la lecture et de l’écriture de données à l’aide des méthodes Document.getSelectedDataAsync, Document.setSelectedDataAsync, Binding.getDataAsync ou Binding.setDataAsync.|
| TextFile  | 1.1 | Word 2013 et versions ultérieures<br>Word 2016 pour Mac<br>Word Online<br>Word pour iPad<br/>|Prise en charge de sortie au format texte (Office.FileType.Text) lors de l’utilisation de la méthode Document.getFileAsync.|

## Méthodes qui ne font pas partie d’un ensemble de conditions requises


Les méthodes suivantes dans l’interface API JavaScript pour Office ne font pas partie d’un ensemble de conditions requises. Si l’une de ces méthodes est nécessaire pour votre complément, utilisez les éléments **Methods** et **Method** dans le manifeste du complément afin de déclarer qu’elles sont obligatoires ou effectuez la vérification à l’exécution à l’aide d’une instruction if. Pour plus d’informations, voir l’article sur la [spécification des conditions requises pour les API et les hôtes Office](../docs/overview/specify-office-hosts-and-api-requirements.md).



|**Nom de la méthode**|**Prise en charge des hôtes Office**|
|:-----|:-----|
|Bindings.addFromPromptAsync|Applications web Access, Excel et Excel Online|
|Document.getFilePropertiesAsync|Excel, Excel Online, Word et PowerPoint|
|Document.getProjectFieldAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getResourceFieldAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getSelectedResourceAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getSelectedTaskAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getSelectedViewAsync|PowerPoint et PowerPoint Online|
|Document.getTaskAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.getTaskFieldAsync|Project Standard 2013 et Project Professionnel 2013|
|Document.goToByIdAsync|Excel, Excel Online, Word et PowerPoint|
|Settings.addHandlerAsync|Applications web Access, Excel, Excel Online, Word et PowerPoint|
|Settings.refreshAsync|Applications web Access, Excel, Excel Online, Word, PowerPoint et PowerPoint Online|
|Settings.removeHandlerAsync|Applications web Access, Excel, Excel Online, Word et PowerPoint|
|TableBinding.clearFormatsAsync|Excel, Excel Online|
|TableBinding.setFormatsAsync|Excel, Excel Online|
|TableBinding.setTableOptionsAsync|Excel, Excel Online|

## Ressources supplémentaires



- [Spécification des exigences en matière d’hôtes Office et d’API](../docs/overview/specify-office-hosts-and-api-requirements.md)

