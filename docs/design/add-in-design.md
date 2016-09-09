# Instructions de conception pour les compléments Office

Les compléments Office prolongent les fonctionnalités d’Office en offrant des fonctions contextuelles auxquelles les utilisateurs peuvent accéder au sein de clients Office. Les compléments permettent aux utilisateurs d’être plus productifs en leur donnant accès à des fonctionnalités tierces au sein d’Office, sans avoir à gérer de coûteux changements de contexte. 

 Votre complément doit s’intégrer de façon harmonieuse avec Office pour fournir une interaction efficace et naturelle à vos utilisateurs. Vous pouvez tirer parti de commandes de complément (extensions de l’interface utilisateur Office) pour permettre aux utilisateurs d’accéder à votre complément, et utiliser les [éléments d’interface utilisateur](ui-elements/ui-elements.md) et les [meilleures pratiques](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices) que nous vous recommandons lorsque vous créez des éléments d’interface utilisateur HTML personnalisés. 
 
 
## Principes fondamentaux de la conception de compléments Office
Quelle que soit l’infrastructure sous-jacente que vous utilisez pour créer votre interface utilisateur personnalisée, appliquez les principes suivants lorsque vous concevez votre complément : 

- **Privilégiez une conception explicitement orientée vers Office**. Les fonctionnalités et l’apparence d’un complément doivent prolonger celles d’Office de façon harmonieuse, notamment en reprenant les thèmes Office ou celui des documents.
 
- **Améliorez l’efficacité des utilisateurs**. Aidez les utilisateurs à mener leurs tâches à bien sans empiéter sur le reste de leur travail. Mettez en œuvre une interaction transparente entre les documents Office et votre complément. 

- **Privilégiez le contenu par rapport aux éléments de détail**. Mettez l’accent sur le contenu et les fonctionnalités du complément plutôt que sur des gadgets accessoires. Optimisez l’utilisation de l’espace en évitant d’ajouter des éléments d’interface superflus qui n’apportent rien à l’expérience utilisateur.  

- **Laissez suffisamment de contrôle aux utilisateurs**. Faites en sorte que les utilisateurs puissent garder le contrôle de ce qu’ils font, qu’ils comprennent les décisions importantes et qu’ils puissent annuler facilement les actions effectuées par le complément. 

- 
  **Prenez en compte toutes les plateformes et les méthodes d’entrée lors de la conception**. Les compléments sont conçus pour fonctionner sur toutes les plateformes prenant en charge Office ; aussi l’expérience utilisateur de votre complément doit-elle être optimisée pour fonctionner avec toutes les plateformes et tous les facteurs de forme. Veillez à ce que votre complément prenne aussi bien en charge les périphériques de type souris/clavier que les appareils et assurez-vous que votre interface utilisateur HTML personnalisée puisse s’adapter à différents facteurs de forme. Pour plus d’informations, consultez notre section relative aux [fonctions tactiles](https://msdn.microsoft.com/EN-US/library/mt590883.aspx#bk_Touch). 


## Langage de création
Nous vous recommandons d’adopter le langage de conception d’Office et d’utiliser la [structure de l’interface utilisateur Office](https://dev.office.com/fabric) pour créer une interface personnalisée utilisant HTML dans vos compléments. Si votre organisation dispose déjà d’un langage de conception, vous pouvez parfaitement utiliser celui-ci, tant que le résultat final offre une expérience harmonieuse pour les utilisateurs d’Office. 


## Blocs de construction pour les compléments
Vous pouvez utiliser deux types d’éléments d’interface utilisateur pour créer vos compléments : 

- Des [commandes de complément](ui-elements/ui-elements.md#add-in-commands), qui vous permettent d’ajouter des hooks natifs à des applications Office.
- Des [éléments d’interface utilisateur HTML personnalisés](ui-elements/ui-elements.md#custom-html-based-ui), qui vous permettent de profiter des avantages du langage HTML dans les clients Office. 

Pour plus d’informations sur l’utilisation de ces blocs de construction, reportez-vous à la rubrique relative aux [éléments d’interface utilisateur](ui-elements/ui-elements.md).  

## Modèles de conception de l’expérience utilisateur

Pour vous aider à créer une expérience utilisateur intéressante pour votre complément, des modèles illustrant les modèles de conception d’expérience utilisateur courants sont disponibles. Ces modèles reflètent les [meilleures pratiques](https://dev.office.com/docs/add-ins/overview/add-in-development-best-practices) pour créer des compléments de qualité attrayants, et comprennent des modèles de conception pour créer des premières expériences d’utilisation, des éléments de personnalisation et des notifications utilisateur. Ils utilisent des composants et des styles de la [structure de l’interface utilisateur Office](https://dev.office.com/fabric) et incluent des éléments qui enrichissent l’interface utilisateur Office.

Pour accéder aux modèles, consultez le référentiel relatif aux [modèles de conception de l’expérience utilisateur du complément Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns). Des fichiers Adobe Illustrator sont également disponibles. Vous pouvez les télécharger et les mettre à jour pour refléter vos propres modèles de conception. Vous pouvez également copier les fichiers de code à partir du référentiel relatif au [code des modèles de conception de l’expérience utilisateur du complément Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) dans votre projet de complément et les personnaliser selon vos besoins. 

## Modèles d’interaction et de mises en page recommandés
Nous fournissons des modèles de mises en page recommandés pour chaque type de complément, avec des exemples **complets** pour vous aider à tout comprendre parfaitement. Pour plus d’informations sur la façon de mettre en page votre complément, reportez-vous aux rubriques suivantes :

- [Mise en page des conteneurs du volet Office](ui-elements/layout-for-task-pane-add-ins.md)
- [Mise en page pour les compléments de contenu](ui-elements/layout-for-content-add-ins.md) 
- [Mises en page pour les compléments de messagerie](ui-elements/layouts-for-outlook-add-ins.md)

Consultez aussi nos modèles d’interaction pour obtenir des exemples de scénarios courants pour les compléments et les modèles d’interaction correspondants.

## Ressources supplémentaires

- [Structure de l’interface utilisateur Office](https://dev.office.com/fabric) 

