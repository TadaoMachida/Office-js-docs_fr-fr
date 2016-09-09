
# Référence de l’interface API JavaScript pour Office

L’interface API JavaScript pour Office vous permet de créer des applications web qui interagissent avec les modèles objet dans les applications hôtes Office. Votre application fera référence à la bibliothèque office.js, qui est un chargeur de script. La bibliothèque office.js charge les modèles objet applicables à l’application Office qui exécute le complément. Vous pouvez utiliser les modèles objet JavaScript suivants :


1. Courant (obligatoire) - API qui ont été introduites avec Office 2013. Il est chargé pour **toutes les applications hôtes Office** et connecte votre application de complément à l’application cliente Office. Le modèle objet contient les API propres aux clients Office et les API applicables à plusieurs applications hôtes clientes Office. Tout le contenu sous [shared](../reference/shared/shared-api.md) et **outlook** correspond aux API courantes. L’espace de noms **Microsoft.Office.WebExtension** (référencé par défaut à l’aide de l’alias [Office](../reference/shared/office.md) dans le code) contient des objets que vous pouvez utiliser pour écrire des scripts qui interagissent avec le contenu dans les documents, feuilles de calcul, présentations, éléments de courrier et projets de vos compléments Office. Vous devez utiliser ces API courantes si votre complément cible Office 2013 et versions ultérieures. Ce modèle objet utilise des rappels.

1. Propre à l’hôte - API qui ont été introduites avec **Office 2016**. Ce modèle objet fournit des objets propres à l’hôte fortement typés qui correspondent aux objets habituels que vous voyez lorsque vous utilisez des clients Office. Il représente l’avenir des API JavaScript Office. Les API propres à l’hôte incluent actuellement l’[API JavaScript Word](../reference/word/word-add-ins-reference-overview.md) et l’[API JavaScript Excel](../reference/excel/application.md). Ce modèle d’objet utilise des promesses.

Sélectionnez le client Office dans la liste déroulante au-dessus de la table des matières pour filtrer le contenu en fonction de votre application hôte cible.

## Applications hôtes prises en charge
* Access
* Excel
* Outlook
* PowerPoint
* Project
* Word

En savoir plus sur les [hôtes pris en charge et les autres exigences](../docs/overview/requirements-for-running-office-add-ins.md)

## Spécifications d’ouverture de l’API

Au fur et à mesure que nous concevons et développons de nouvelles API pour les compléments Office, nous les mettons à votre disposition sur notre page de [spécifications d’ouverture de l’API](openspec.md) pour que vous puissiez fournir vos commentaires. Découvrez les nouvelles fonctionnalités dans le pipeline et donnez votre avis sur nos spécifications de conception.

