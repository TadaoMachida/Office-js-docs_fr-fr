
# Déployer et installer des compléments Outlook à des fins de test


Dans le cadre du processus de développement d’un complément Outlook, vous devrez déployer et installer de façon itérative le complément à des fins de test, ce qui implique les étapes suivantes :


1. Création d’un fichier manifeste qui décrit le complément.
    
2. Déploiement du ou des fichiers de l’interface utilisateur du complément sur un serveur web.
    
3. Installation du complément dans votre boîte aux lettres.
    
4. Test du complément, mise en œuvre des modifications appropriées dans l’interface utilisateur ou dans les fichiers manifeste, et répétition des étapes 2 et 3 pour tester les modifications.
    

## Création d’un fichier manifeste pour le complément

Chaque complément est décrit par un manifeste XML, un document qui fournit au serveur des informations sur le complément, décrit le complément pour l’utilisateur et identifie l’emplacement du fichier HTML de l’interface utilisateur du complément. Vous pouvez stocker le manifeste dans un dossier local ou sur un serveur, à condition que le complément soit accessible par le serveur Exchange de la boîte aux lettres avec laquelle vous procédez aux tests. Nous partons du principe que vous stockez votre manifeste dans un dossier local. Pour plus d’informations sur la création d’un fichier manifeste, voir [Manifestes des compléments Outlook](../outlook/manifests/manifests.md). 


## Déploiement d’un complément sur un serveur web

Vous pouvez créer l’interface utilisateur du complément en HTML et en JavaScript. Le fichier source obtenu est stocké sur un serveur web auquel a accès le serveur Exchange qui héberge le complément. Ce fichier source est identifié par l’élément enfant  **SourceLocation** dans les éléments [DesktopSettings](http://msdn.microsoft.com/en-us/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c%28Office.15%29.aspx), [TabletSettings](http://msdn.microsoft.com/en-us/library/5c89cc7c-7ae0-49c9-fdd5-4c52118228f6%28Office.15%29.aspx) et/ou [PhoneSettings](http://msdn.microsoft.com/en-us/library/13e4eae3-8e8c-fd55-a1c2-3297b485f327%28Office.15%29.aspx) spécifiés dans le fichier manifeste du complément.

Après le déploiement initial des fichiers d’interface utilisateur pour le complément, vous pouvez mettre à jour l’interface utilisateur et le comportement du complément en remplaçant le fichier HTML stocké sur le serveur web par une nouvelle version du fichier HTML.


## Installation du complément


Après la préparation du fichier manifeste du complément et le déploiement de son interface utilisateur sur un serveur web accessible, vous pouvez installer le complément pour une boîte aux lettres sur un serveur Exchange à l’aide d’un client riche Outlook, d’Outlook Web App ou d’OWA pour périphériques, ou en exécutant des applets de commande Windows PowerShell à distance.


### Installation d’un complément dans un client riche Outlook

Vous pouvez installer un complément si votre boîte aux lettres se trouve sur Exchange Online, Exchange 2013 ou version ultérieure. Dans Outlook pour Windows, vous pouvez installer des compléments via le mode Backstage d’Office Fluent. Sélectionnez **Fichier** et **Gérer les compléments**. Cela vous permet de vous connecter au Centre d’administration Exchange. Une fois connecté, continuez le processus d’installation à l’étape 4 de la section suivante.

Dans Outlook pour Mac, choisissez **Gérer les compléments** à l’extrémité droite de la barre des compléments, puis connectez-vous au Centre d’administration Exchange. Passez à l’étape 4 de la section suivante.


### Installation d’un complément à l’aide d’Outlook Web App ou d’Outlook.com

Pour utiliser Outlook Web App (OWA) pour installer un complément Outlook, procédez comme suit :


1. Accédez à l’URL OWA pour votre organisation ou à Outlook.com et connectez-vous.
    
2. Sélectionnez l’icône d’engrenage dans le coin supérieur droit et choisissez **Gérer les compléments**.
    
3. Sélectionnez le signe plus ( **+**) pour ajouter un nouveau complément.
    
4. Dans la liste déroulante, sélectionnez **Ajouter à partir d’un fichier**, en partant du principe que vous avez stocké le manifeste dans un dossier local.
    
5. Accédez au chemin du manifeste, puis sélectionnez **Installer**.
    
6. Sélectionnez le nom d’utilisateur dans le coin supérieur droit de la fenêtre et sélectionnez **Mon courrier** pour accéder à votre message électronique afin de tester le complément.
    

>**Remarque :**  si vous n’utilisez aucun des éléments suivants pour développer votre complément : 
- Client de développeur Office 365
- Outils de développement Office 365 Napa
- Visual Studio

Et, si vous n’avez pas au minimum le rôle « Mes applications personnalisées » de votre serveur Exchange, vous pouvez installer les compléments uniquement à partir de l’Office Store. Pour tester votre complément ou installer des compléments en général en spécifiant une URL ou un nom de fichier pour le manifeste de complément, vous devez demander à votre administrateur Exchange de vous octroyer les autorisations nécessaires.

L’administrateur Exchange peut exécuter la cmdlet PowerShell suivante pour affecter les autorisations nécessaires à un seul utilisateur. Dans cet exemple, « wendyri » est l’alias de messagerie de l’utilisateur.

```New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"```

Selon les besoins, l’administrateur peut exécuter l’applet de commande suivante pour affecter des autorisations nécessaires similaires à plusieurs utilisateurs :

```$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}```

Pour plus d’informations sur le rôle « Mes compléments personnalisés », consultez la rubrique relative au [rôle « Mes compléments personnalisés »](http://technet.microsoft.com/en-us/library/aa0321b3-2ec0-4694-875b-7a93d3d99089%28EXCHG.150%29.aspx). 

L’utilisation d’Office 365, des Outils de développement Office 365 « Napa » ou de Visual Studio pour développer des compléments vous amène à endosser le rôle d’administrateur d’organisation, ce qui vous permet d’installer des compléments par fichier ou par URL dans le Centre d’administration Exchange ou via des cmdlets PowerShell.


### Installation d’un complément à l’aide de PowerShell à distance

Après avoir créé une session Windows PowerShell à distance sur votre serveur Exchange, vous pouvez installer un complément Outlook en utilisant l’applet de commande  **New-App** avec la commande PowerShell suivante.


```
New-App -URL:"http://<fully-qualified URL">
```

L’URL complète est l’emplacement du fichier de manifeste de complément que vous avez préparé pour votre complément.

Vous pouvez utiliser les applets de commande supplémentaires suivantes pour gérer les compléments pour une boîte aux lettres :


-  **Get-App** - répertorie les compléments activés pour une boîte aux lettres.
    
-  **Set-App** - active ou désactive un complément sur une boîte aux lettres.
    
-  **Remove-App** - supprime un complément précédemment installé à partir d’un serveur Exchange.
    

## Ressources supplémentaires



- [Compléments Outlook](../outlook/outlook-add-ins.md)
    
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)
    
