
# Configuration requise pour exécuter des compléments Office


Cet article décrit la configuration logicielle et matérielle requise pour l’exécution des compléments Office.

>**Remarque :** Pour savoir de manière détaillée quelle version d’Office prend en charge les compléments Office, consultez la page relative à la [disponibilité des compléments Office sur les plateformes et les hôtes](http://dev.office.com/add-in-availability). 


## Exigences en matière de serveur

Pour pouvoir installer et exécuter des Complément Office, vous devez d’abord déployer les fichiers manifeste et de pages web pour l’interface utilisateur et le code de votre complément sur les emplacements de serveur appropriés.

Pour tous les types de complément (compléments de contenu, Outlook et volet Office, et les commandes de compléments), vous devez déployer les fichiers de pages web de votre complément sur un serveur web ou un service d’hébergement web, tel que [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).


 >**Remarque :**   lorsque vous développez et déboguez un complément dans Visual Studio, Visual Studio déploie et exécute les fichiers de pages web de votre complément localement avec IIS Express, et ne requiert pas de serveur web supplémentaire. De la même manière, lorsque vous développez et déboguez un complément avec les Outils de développement Office 365 « Napa » dans le navigateur, il déploie et exécute les fichiers de pages web de votre complément à partir de l’emplacement de stockage associé au compte que vous avez utilisé pour vous connecter à Napa.

Pour les compléments du volet Office et de contenu, dans les applications hôtes Office prises en charge (applications web Access, Word, Excel, PowerPoint ou Project), vous avez également besoin d’un [catalogue de compléments](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) sur SharePoint pour télécharger les fichiers manifeste XML des compléments.

Pour tester et exécuter des compléments Outlook, le compte de messagerie Outlook de l’utilisateur doit être situé sur Exchange 2013 ou une version ultérieure, disponible par le biais d’Office 365, Exchange Online ou via une installation sur site. L’utilisateur ou l’administrateur installe les fichiers manifeste pour les compléments Outlook sur ce serveur.

 >**Remarque :**   Les comptes de messagerie POP et IMAP dans Outlook ne prennent pas en charge les Compléments Office.




## Configuration requise pour le client : Ordinateur de bureau et tablette Windows

Le logiciel suivant est requis pour développer un Complément Office pour les clients Office ou les clients web pris en charge qui s’exécutent sur un ordinateur de bureau, un ordinateur portable ou une tablette Windows :


- Pour les ordinateurs de bureau Windows x86 et x64 et les tablettes telles que Surface Pro :

    - La version 32 bits ou 64 bits d’Office 2013 ou une version ultérieure s’exécutant sur Windows 7 ou une version ultérieure.

    - Excel 2013, Outlook 2013, PowerPoint 2013, Project Professionnel 2013, Project 2013 SP1 ou Word 2013, ou une version ultérieure du client Office, si vous testez ou exécutez un Complément Office, notamment pour l’un de ces clients de bureau Office. Les clients de bureau Office peuvent être installés sur site ou par le biais de « Démarrer en un clic » sur l’ordinateur client.

- Internet Explorer 9 ou une version ultérieure, qui doit être installé mais pas nécessairement défini comme le navigateur par défaut. Pour prendre en charge des Compléments Office, le client Office qui sert d’hôte utilise des composants de navigateur faisant partie d’Internet Explorer 9 ou d’une version ultérieure.

- L’un des navigateurs suivants comme navigateur par défaut : Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13 ou une version ultérieure de ces navigateurs.

- Un éditeur HTML et JavaScript tel que le Bloc-notes, [Visual Studio et les outils de développement Office ](https://www.visualstudio.com/features/office-tools-vs) ou un outil de développement web tiers.


## Exigences en matière de client : ordinateur de bureau OS X

Outlook pour Mac, qui est distribué dans le cadre d’Office 365, prend en charge les compléments Outlook. L’exécution des compléments Outlook sur Outlook pour Mac a les mêmes exigences qu’Outlook pour Mac lui-même : le système d’exploitation doit être au minimum OS X v10.10 « Yosemite ». Comme Outlook pour Mac utilise WebKit comme moteur de disposition pour restituer les pages de complément, il n’existe pas de dépendance de navigateur supplémentaire.

Les versions de client minimales d’Office pour Mac prenant en charge les compléments Office sont les suivantes :
- Word pour Mac version 15.18 (160109) 
- Excel pour Mac version 15.19 (160206) 
- PowerPoint pour Mac version 15.24 (160614)

## Configuration requise pour le client : Prise en charge du navigateur pour les clients web Office Online et SharePoint

Tout navigateur qui prend en charge ECMAScript 5.1, HTML5 et CSS3, tel qu’Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6 ou une version ultérieure de ces navigateurs.


## Exigences en matière de client : smartphone et tablette autres que Windows

Plus particulièrement pour OWA pour périphériques et Outlook Web App exécutés dans un navigateur sur des smartphones et des tablettes non Windows, le logiciel suivant est requis pour tester et exécuter des compléments Outlook.


| Application hôte | Appareil | Système d’exploitation | Compte Exchange | Navigateur mobile |
|:-----|:-----|:-----|:-----|:-----|
|OWA pour Android|Smartphones Android. D’un point de vue technique, ces appareils sont considérés comme « petits » ou « normaux » par [Android OS](https://developer.android.com/guide/practices/screens_support.html).|Android 4.4 Kitkat ou version ultérieure|Sur la dernière mise à jour d’Office 365 pour les entreprises ou d’Exchange Online|Complément natif pour Android, navigateur non applicable|
|OWA pour iPad|iPad 2 ou version ultérieure|iOS 6 ou version ultérieure|Sur la dernière mise à jour d’Office 365 pour les entreprises ou d’Exchange Online|Complément natif pour iOS, navigateur non applicable|
|OWA pour iPhone|iPhone 4S ou version ultérieure|iOS 6 ou version ultérieure|Sur la dernière mise à jour d’Office 365 pour les entreprises ou d’Exchange Online|Complément natif pour iOS, navigateur non applicable|
|Outlook Web App|iPhone 4, iPad 2, iPod Touch 4 (ou version ultérieure de ces appareils)|iOS 5 ou version ultérieure|Sur Office 365, Exchange Online, ou localement sur Exchange Server 2013 ou version ultérieure|Safari|


## Ressources supplémentaires

- [Vue d’ensemble de la plateforme des compléments pour Office](../../docs/overview/office-add-ins.md)
- [Disponibilité des compléments Office sur les plateformes et les hôtes](http://dev.office.com/add-in-availability)

