
# Conception de compléments Office pour l’iPad


Le tableau suivant répertorie les tâches à effectuer pour concevoir un complément Office à exécuter dans Office pour iPad.


|**Task**|**Description**|**Resources**|
|:-----|:-----|:-----|
|Mettez à jour votre complément pour prendre en charge la version 1.1 d’Office.js.|Mettez à jour les fichiers JavaScript (Office.js et fichiers .js propres aux applications) et le fichier de validation du manifeste du complément utilisés dans votre projet Complément Office vers la version 1.1.|[Nouveautés de l’API JavaScript pour Office](../../reference/what's-changed-in-the-javascript-api-for-office.md)|
|Appliquez les méthodes recommandées pour concevoir une interface utilisateur.|Intégrez l’interface utilisateur de votre complément de manière transparente avec iOS.|[Concevoir pour iOS](https://developer.apple.com/library/ios/documentation/UserExperience/Conceptual/MobileHIG/)|
|Appliquez les méthodes recommandées pour concevoir un complément.|Assurez-vous que votre complément offre une valeur claire, une expérience conviviale et des performances optimales.|[Meilleures pratiques en matière de développement de compléments Office](../../docs/overview/add-in-development-best-practices.md)|
|Optimisez votre complément pour les écrans tactiles.|Concevez une interface utilisateur optimisée pour les écrans tactiles, en plus de la souris et du clavier.|[Application des principes de conception de l’expérience utilisateur](https://msdn.microsoft.com/fr-fr/library/mt590883.aspx#Anchor_3)|
|Proposez un complément gratuit.|Office pour iPad vous permet de communiquer avec davantage d’utilisateurs et de promouvoir vos services. Ces nouveaux utilisateurs peuvent devenir vos clients.|[Stratégie de validation 10.8](http://msdn.microsoft.com/fr-fr/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|Proposez un commerce de complément gratuit.|Votre complément ne doit pas comporter de services payants, d’offres d’essai, une interface utilisateur destinée à inciter à la vente, ni de liens vers des magasins en ligne où les utilisateurs peuvent acheter ou acquérir d’autres contenus, applications ou compléments. Vos pages Politique de confidentialité et Conditions d’utilisation ne doivent pas non plus comporter de liens vers une interface utilisateur commerciale ou le Store.|[Stratégie de validation 3.4](http://msdn.microsoft.com/fr-fr/library/cd90836a-523e-42f5-ab02-5123cdf9fefe%28Office.15%29.aspx)|
|Renvoyez votre complément à l’Office Store.|Dans le tableau de bord vendeur, cochez la case **Rendre ce complément accessible dans le catalogue de compléments Office sur iPad**. Indiquez votre ID de développeur Apple dans la case Identifiant Apple. Lisez le [Contrat du fournisseur d’application Office Store](https://sellerdashboard.microsoft.com/Assets/Content/Agreements/en-US/Office_Store_Seller_Agreement_20120927.md) pour connaître les termes du contrat.|[Soumission des compléments SharePoint et Office, ainsi que des applications web Office 365 dans l’Office Store](http://msdn.microsoft.com/fr-fr/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)|
Votre complément peut rester en l’état pour les applications Office exécutées sur d’autres plateformes. Vous pouvez également proposer une interface utilisateur différente en fonction du navigateur ou de l’appareil qui utilise votre complément. Pour savoir si votre complément est exécuté sur un iPad, vous pouvez utiliser les API suivantes : 

- var isTouchEnabled = [Office.context.touchEnabled](../../reference/shared/office.context.touchenabled.md)
    
- var allowCommerce = [Office.context.commerceAllowed](../../reference/shared/office.context.commerceallowed.md)
    

## Meilleures pratiques en matière de développement de compléments Office pour iOS et Mac

Appliquez les meilleures pratiques suivantes pour développer des compléments pour iOS :


-  **Utilisez Visual Studio pour développer votre complément.**
    
    If you develop your add-in with Visual Studio, you can [set breakpoints and debug its code](../get-started/create-and-debug-office-add-ins-in-visual-studio.md#Test) in an Office host application running on Windows, before sideloading your add-in on the iPad or Mac. Because an add-in that runs in Office for iOS or Office for Mac supports the same APIs as an add-in running in Office for Windows, your add-in's code should run the same way on both platforms.
    
-  **Précisez les conditions de fonctionnement de l’API dans le manifeste de votre complément ou avec des vérifications à l’exécution.**
    
    When you specify API requirements in your add-in's manifest, Office will determine if the host application supports those API members. If the API members are available in the host, then your add-in will be available in that host application. Alternatively, you can perform a runtime check to determine if a method is available in the host before using it in your add-in. Runtime checks ensure that your add-in is always available in the host, and provides additional functionality if the methods are available. For more information, see [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md).
    
Pour plus d’informations sur des pratiques plus générales en matière de développement de compléments, consultez la rubrique [Meilleures pratiques en matière de développement de compléments Office](../../docs/overview/add-in-development-best-practices.md).


## Ressources supplémentaires
<a name="bk_addresources"></a>


- [Charger une version test d’un complément Office sur iPad ou Mac](../../docs/testing/sideload-an-office-add-in-on-ipad-and-mac.md)
    
- [Débogage des compléments Office sur iPad et Mac](../../docs/testing/debug-office-add-ins-on-ipad-and-mac.md)
    

