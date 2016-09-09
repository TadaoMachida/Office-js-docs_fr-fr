
# Mettre à jour la version de votre API JavaScript pour Office et des fichiers de schéma de manifeste



Cet article décrit comment mettre à jour vers la version 1.1 les fichiers JavaScript pour Office (Office.js et fichiers .js propres aux applications) et le fichier de validation du manifeste du complément utilisés dans votre projet de complément Office.

## Utilisation des fichiers de projet les plus récents

Si vous utilisez Visual Studio pour développer votre complément, et que vous souhaitez utiliser les [nouveaux membres d’API](../../reference/what's-changed-in-the-javascript-api-for-office.md) de l’interface API JavaScript pour Office et les [fonctionnalités de la version 1.1 du manifeste du complément](../../docs/overview/add-in-manifests.md) (qui est validé par rapport à offappmanifest-1.1.xsd), vous devez télécharger et installer [Visual Studio 2015 et les derniers Outils de développement Office](https://www.visualstudio.com/features/office-tools-vs).

Si vous utilisez un éditeur de texte ou une interface IDE autre que Visual Studio pour développer votre complément, vous devez mettre à jour les références vers le CDN pour Office.js et la version de schéma référencée dans le manifeste de votre complément.

Pour exécuter un complément développé à l’aide des fonctionnalités nouvelles et mises à jour du manifeste du complément et de l’API d’Office.js, vos clients doivent exécuter des produits locaux Office 2013 SP1 ou version ultérieure, et le cas échéant, SharePoint Server 2013 SP1 et des produits serveur associés, Exchange Server 2013 Service Pack 1 (SP1) ou des produits hébergés en ligne équivalents : Office 365, SharePoint Online et Exchange Online.

Pour télécharger des produits Office, SharePoint et Exchange SP1, voir :


- [Liste de toutes les mises à jour Service Pack 1 (SP1) pour Microsoft Office 2013 et les produits bureautiques connexes](http://support.microsoft.com/kb/2850036)
    
- [Liste de toutes les mises à jour Service Pack 1 (SP1) pour Microsoft SharePoint Server 2013 et les produits serveur connexes](http://support.microsoft.com/kb/2850035)
    
- [Description du Service Pack 1 d’Exchange Server 2013](http://support.microsoft.com/kb/2926248)
    

## Mise à jour d’un projet de Complément Office créé à l’aide de Visual Studio pour utiliser le schéma de manifeste de complément version 1.1 et la bibliothèque de l’API JavaScript pour Office la plus récente


Pour les projets créés avant la sortie de la version 1.1 de l’API JavaScript pour Office et du schéma de manifeste de complément, vous pouvez mettre à jour les fichiers d’un projet à l’aide du  **gestionnaire de package NuGet**, puis mettre à jour les pages HTML de votre complément pour les référencer. 

Notez que le processus de mise à jour est appliqué  _par projet_  ; vous devrez répéter le processus de mise à jour pour chaque projet de complément dans lequel vous souhaitez utiliser la version 1.1 d’Office.js et du schéma de manifeste de complément.




### Mise à jour des fichiers de bibliothèque de l’API JavaScript pour Office dans votre projet vers la dernière version


1. Dans Visual Studio 2015, ouvrez ou créez un projet **Complément Office**.
    
      - Dans le volet de gauche, sélectionnez **Mettre à jour** et terminez le processus de mise à jour du package.
    
  - Passez à l’étape 6.
    
2. Sélectionnez **Outils**  >  **Gestionnaire de package NuGet**  >  **Gérer les packages NuGet pour la solution**.
    
3. Dans **Gestionnaire de package NuGet**, sélectionnez  **nuget.org** pour **Source du package** et **Mise à niveau disponible** pour **Filtrer**, puis sélectionnez Microsoft.Office.js.
    
4. Dans le volet de gauche, sélectionnez **Mettre à jour** et terminez le processus de mise à jour du package.
    
5. Dans la balise  **head** des pages HTML de votre complément, commentez ou supprimez toute référence de script office.js existante (par exemple :`<script src="https://appsforoffice.microsoft.com/lib/1.0/hosted/office.js" type="text/javascript"></script>)`), puis référencez la bibliothèque de l’interface API JavaScript pour Office mise à jour de cette façon (en remplaçant la valeur de version par  « 1 »). 

   >**Remarque**La valeur  « /1/ » devant office.js dans l’URL CDN ci-dessous indique d’utiliser la dernière version incrémentielle au sein de la version 1 d’Office.js.
    
```
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```


### Pour mettre à jour le fichier manifeste dans votre projet afin d’utiliser la version 1.1 du schéma


- Dans le fichier de manifeste de complément de votre projet ( _projectname_ Manifest.xml), mettez à jour l’attribut **xmlns** de l’élément **OfficeApp** en appliquant la valeur « 1.1 » à la version (sans modifier les attributs autres que **xmlns**) :
    
```XML
  <OfficeApp xsi:type="ContentApp" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" >
```


>
  **Remarque**  Après la mise à jour de la version du schéma de manifeste de complément vers 1.1, vous devrez supprimer les éléments **Capabilities** et **Capability**, et les remplacer par les [éléments Hosts et Host](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx) ou les [éléments Requirements et Requirement](../../docs/overview/specify-office-hosts-and-api-requirements.md).

## Mise à jour d’un projet d’Complément Office créée à l’aide d’un éditeur de texte ou d’une autre interface IDE en vue d’utiliser le schéma de manifeste de complément version 1.1 ou la dernière bibliothèque de l’API JavaScript pour Office


Pour les projets créés avant la publication de la version 1.1 de l’API JavaScript pour Office et du schéma de manifeste de complément, vous devez mettre à jour les pages HTML de votre complément afin de faire référence au CDN de la bibliothèque version 1.1, ainsi que mettre à jour le fichier de manifeste de votre complément pour utiliser le schéma version 1.1. 

Le processus de mise à jour est appliqué  _par projet_  ; vous devrez répéter le processus de mise à jour pour chaque projet de complément dans lequel vous souhaitez utiliser la version 1.1 d’Office.js et du schéma de manifeste de complément.

Vous n’avez pas besoin de copies locales des fichiers de l’interface API JavaScript pour Office (fichiers Office.js et fichiers .js propres aux applications) pour développer un complément Office (si vous référencez le CDN pour Office.js, les fichiers requis sont téléchargés lors de l’exécution), mais si vous voulez une copie locale des fichiers de bibliothèque, vous pouvez utiliser l’[utilitaire de ligne de commande NuGet](http://docs.nuget.org/consume/installing-nuget) et la commande `Install-Package Microsoft.Office.js` pour les télécharger.

 > **Remarque** : pour obtenir une copie du fichier XSD (définition de schéma XML) pour le manifeste de complément version 1.1, consultez la liste dans [Informations de référence sur le schéma des manifestes des applications pour Office (version 1.1)](../overview/add-in-manifests.md).


### Mise à jour des fichiers de bibliothèque de l’API JavaScript pour Office dans votre projet pour utiliser la dernière version


1. Ouvrez les pages HTML de votre complément dans un éditeur de texte ou une interface IDE.
    
2. Dans la balise **head** des pages HTML de votre complément, commentez ou supprimez toute référence de script office.js existante (par exemple : `<script src="https://appsforoffice.microsoft.com/lib/1.0/hosted/office.js" type="text/javascript"></script>)`), puis référencez la bibliothèque de l’interface API JavaScript pour Office mise à jour de cette façon (en remplaçant la valeur de version par  « 1 »).
    
```
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```


    The  `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.
    

### Pour mettre à jour le fichier manifeste dans votre projet afin d’utiliser la version 1.1 du schéma


- Dans le fichier de manifeste de complément de votre projet ( _projectname_ Manifest.xml), mettez à jour l’attribut **xmlns** de l’élément **OfficeApp** en appliquant la valeur `1.1` à la version (sans modifier les attributs autres que **xmlns**) :
    
```XML
<OfficeApp xsi:type="ContentApp" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" >
```

>
  **Remarque**  Après la mise à jour de la version du schéma de manifeste de complément vers 1.1, vous devrez supprimer les éléments **Capabilities** et **Capability**, et les remplacer par les [éléments Hosts et Host](http://msdn.microsoft.com/library/cff9fbdf-a530-4f6e-91ca-81bcacd90dcd%28Office.15%29.aspx) ou les [éléments Requirements et Requirement](../../docs/overview/specify-office-hosts-and-api-requirements.md).
    

## Ressources supplémentaires



- [Spécification des exigences en matière d’hôtes Office et d’API](../../docs/overview/specify-office-hosts-and-api-requirements.md)
    
- [Présentation de l’API JavaScript pour Office](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Interface API JavaScript pour Office](../../reference/javascript-api-for-office.md)
    
- [Informations de référence sur le schéma des manifestes des applications pour Office (version 1.1)](../overview/add-in-manifests.md)
    
