
# Authentifier un complément Outlook à l’aide de jetons d’identité Exchange

Votre complément Outlook peut fournir à vos clients des informations de toute provenance sur Internet, que ce soit à partir du serveur qui héberge le complément, de votre réseau interne ou de tout autre endroit sur le cloud. Cependant, si ces informations sont protégées, votre complément Outlook doit pouvoir associer le compte de messagerie Exchange à votre service d’information. Exchange 2013 peut activer l’authentification unique pour votre complément en fournissant un jeton qui identifie le compte de messagerie à l’origine de la demande. Vous pouvez associer ce jeton à un utilisateur inscrit auprès de votre complément afin qu’il soit reconnu dès lors que le complément se connecte à votre service.

## Jetons d’identité


Deux de nos exemples de compléments utilisent des informations publiquement disponibles, une première affiche une carte Bing pour des adresses dans un message, une autre affiche un aperçu pour vos liens vidéo YouTube dans un message. Mais votre complément peut également accéder à des informations non publiques. Vous pouvez utiliser le serveur qui héberge votre complément pour lier ce dernier aux informations de votre réseau interne ou se trouvant ailleurs dans le cloud.

Vous pouvez utiliser de nombreuses techniques différentes pour identifier et authentifier les utilisateurs d’un complément. Exchange 2013 simplifie l’authentification de l’utilisateur en fournissant à votre complément un jeton d’identité qui identifie un compte de messagerie Exchange spécifique. Vous pouvez associer ce jeton dans votre service à un utilisateur inscrit, ce qui active l’authentification unique de vos clients utilisant des compléments Outlook. 

Pour utiliser l’authentification unique dans votre complément, le code procède ainsi :


* Appelle une fonction dans l’API du complément Outlook qui renvoie un jeton d’identité.
* Envoie le jeton avec une demande à votre serveur.
* Décompresse la réponse du serveur pour afficher les informations provenant de votre service.
    
Côté serveur, les choses sont relativement plus complexes. Lorsque votre serveur reçoit une demande de votre complément Outlook, le processus opère de la façon suivante :

* Le serveur valide le jeton. Vous pouvez utiliser notre [bibliothèque de validation de jeton gérée](../../docs/outlook/use-the-token-validation-library.md) ou [créer votre propre bibliothèque](../../docs/outlook/validate-an-identity-token.md) pour votre service.
* Le serveur recherche l’identificateur unique dans le jeton pour voir s’il est associé à une identité connue. Votre service doit [implémenter une méthode qui fait correspondre l’identificateur](../../docs/outlook/authenticate-a-user-with-an-identity-token.md) à des utilisateurs connus de votre service.
* Si l’identificateur unique correspond à un identificateur précédemment stocké avec un ensemble d’informations d’identification sur le serveur, votre serveur peut répondre en fournissant les informations demandées sans que le client n’ait besoin de se connecter à votre service.
* Si l’identificateur unique est inconnu, le serveur envoie une réponse demandant à l’utilisateur de se connecter avec des informations d’identification pour le serveur.
* Si les informations d’identification correspondent à une identité connue sur le serveur, vous pouvez faire correspondre cette identité à l’identificateur unique dans le jeton afin que lors d’une prochaine demande, votre serveur puisse répondre sans requérir une étape de connexion supplémentaire.

 >**Remarque**  Ceci n’est qu’une simple suggestion d’utilisation du jeton d’identité. Comme toujours, lorsque vous traitez des informations d’identité et d’authentification, vous devez vous assurer que le code répond aux exigences de sécurité de votre organisation.

Examinons les spécificités de la fonction. Dans les articles suivants, nous utiliserons un complément Outlook simple qui envoie à un service web le jeton d’identité et une liste de numéros de téléphone trouvés dans le message. 

- [Présentation du jeton d’identité Exchange](../outlook/inside-the-identity-token.md)
- [Appeler un service à partir d’un complément Outlook à l’aide d’un jeton d’identité dans Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
- [Utiliser la bibliothèque de validation des jetons Exchange](../outlvalidate-an-identity-token.md ook/use-the-token-validation-library.md)
- [Valider un jeton d’identité Exchange](../outlook/validate-an-identity-token.md )
- [Authentifier un utilisateur avec un jeton d’identité pour Exchange](../outlook/validate-an-identity-token.md)


## Ressources supplémentaires



- [Compléments Outlook](../outlook/outlook-add-ins.md)
    
- [Appeler des services web à partir d’un complément Outlook](../outlook/web-services.md)
    


