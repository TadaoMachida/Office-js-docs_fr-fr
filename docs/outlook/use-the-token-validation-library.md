
# Utilisation de la bibliothèque de validation de l’API managée Exchange Web Services

Vous pouvez identifier les clients de votre complément Outlook à l’aide d’un jeton d’identité que votre complément demande à un serveur exécutant Exchange Server 2013 ou Exchange Online. Ce jeton, formaté comme un jeton Web JSON, fournit un identificateur unique pour un compte de messagerie sur un serveur Exchange Server. L’API managée des services web Exchange (EWS) fournit des classes d’assistance pour simplifier l’utilisation du jeton d’identité.

## Conditions préalables à l’utilisation de la bibliothèque de validation

Pour valider un jeton d’identité Exchange, vous devez installer la [bibliothèque de l’API managée EWS](https://www.nuget.org/packages/Microsoft.Exchange.WebServices).

## Validation du jeton d’identité Exchange

La bibliothèque de validation de l’API managée EWS fournit la classe  **AppIdentityToken** pour gérer les jetons d’identité Exchange. La méthode suivante montre comment créer une instance **AppIdentityToken** et appeler la méthode **Validate** pour vérifier la validité du jeton. La méthode tient compte des paramètres suivants :

- *rawToken* : Représentation de la chaîne du jeton renvoyé dans votre complément Outlook à partir de la méthode [**Office.context.mailbox.getUserIdentityTokenAsync**](http://dev.office.com/reference/add-ins/outlook/Office.context.mailbox).
- *hostUri* : URI complète de la page dans votre complément Outlook ayant appelé la méthode **getUserIdentityTokenAsync**.

```C#
// Required to use the validation library.
using Microsoft.Exchange.WebServices.Auth.Validate;

private AppIdentityToken CreateAndValidateIdentityToken(string rawToken, string hostUri)
{
    try
    {
        AppIdentityToken token = (AppIdentityToken)AuthToken.Parse(rawToken);
        token.Validate(new Uri(hostUri));

        return token;
    }
    catch (TokenValidationException ex)
    {
        throw new ApplicationException("A client identity token validation error occurred.", ex);
    }
}
```

## Ressources supplémentaires

- [Authentifier un complément Outlook à l’aide de jetons d’identité Exchange](../outlook/authentication.md)  
- [Présentation du jeton d’identité Exchange](../outlook/inside-the-identity-token.md)
- [Valider un jeton d’identité Exchange](../outlook/validate-an-identity-token.md)
    
