
# Authentifier un utilisateur avec un jeton d’identité pour Exchange

Vous pouvez implémenter un schéma d’authentification unique pour un service d’information qui permet à vos clients utilisant des compléments Outlook de se connecter à votre service à l’aide des informations d’identification de leur serveur Exchange. Cet article montre comment faire correspondre des informations d’identification à l’aide d’un simple magasin de données utilisateur basé sur un objet  **Dictionary**.

 >**Remarque**  Il s’agit d’un simple exemple de l’authentification unique que vous ne devez pas utiliser dans votre code de production. Comme toujours, lorsqu’il est question d’identité et d’authentification, vous devez vous assurer que votre code respecte les exigences en matière de sécurité de votre organisation.


## Conditions préalables à l’utilisation de l’authentification unique


Pour utiliser un jeton de sécurité pour l’authentification unique, votre application de service a besoin d’un jeton d’identité valide. Pour plus d’informations sur les jetons d’identité et la manière de demander et valider un jeton d’identité, voir les articles suivants :


- [Présentation du jeton d’identité Exchange](../outlook/inside-the-identity-token.md)
    
- [Appeler un service à partir d’un complément Outlook à l’aide d’un jeton d’identité dans Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [Utiliser la bibliothèque de validation des jetons Exchange](../outlook/use-the-token-validation-library.md) si vous utilisez du code managé ou [Valider un jeton d’identité Exchange](../outlook/validate-an-identity-token.md) si vous écrivez votre propre méthode de validation de jeton.
    

## Authentifier un utilisateur


L’exemple de code suivant montre un objet d’authentification simple qui fait correspondre l’identité unique représentée par un jeton d’identité à un ensemble d’informations d’identification pour un service. La classe  **TokenAuthentication** fournit une méthode, **GetResponseFromService**, qui retourne une réponse pour les jetons précédemment authentifiés, ou qui invite l’utilisateur à fournir des informations d’identification pouvant être authentifiées et associées au jeton d’identité. Le code n’est pas complet ; il suppose que vous fournissiez les objets et méthodes suivants.



|**Objet/Méthode**|**Description**|
|:-----|:-----|
|Objet **LocalCredentials**|Représente les informations d’identification de l’utilisateur de votre service. La structure de l’objet dépend des exigences de votre service.|
|Objet **IdentityToken**|Contient un jeton d’identité utilisateur envoyé à votre service par un complément Outlook. L’objet doit contenir au moins l’identificateur Exchange unique de l’utilisateur et l’URL des métadonnées d’authentification du serveur qui a émis le jeton. Cet exemple utilise l’objet jeton d’identité défini dans l’article [Valider un jeton d’identité Exchange](../outlook/validate-an-identity-token.md).|
|Objet **JsonResponse**|Représente la réponse donnée par votre service. L’objet peut être sérialisé en objet JSON.|
|Méthode **CallService**|Appelle votre service avec un objet  **LocalCredentials** qui contient les informations d’identification de l’utilisateur du service et un objet qui contient les données de la demande de service. Si les informations d’identification sont valides, cette méthode renvoie un objet **JsonReponse** qui contient les résultats de la demande. Si elles ne sont pas valides, cette méthode renvoie **null**.|
|Méthode **GetCredentialsResponse**|Renvoie un objet  **JsonReponse** que votre complément de messagerie pour Office reconnaît en tant que demande d’informations d’identification pour le service.|
|Méthode **LocalCredentialsAreValid**|Renvoie  **true** si les informations d’identification fournies au service sont valides ; sinon, renvoie **false**.|

 >**Remarque**  Il s’agit simplement d’une suggestion d’utilisation du jeton d’identité. Comme toujours, lorsqu’il est question d’identité et d’authentification, vous devez vous assurer que votre code respecte les exigences en matière de sécurité de votre organisation.


```C#
    public class TokenAuthentication
    {
        // This example uses a Dictionary object to store local credentials. Your application should use
        // a data store that is appropriate to the security requirements of your organization.
        private Dictionary<string, LocalCredentials> AuthenticationCache = new Dictionary<string, LocalCredentials>();

        // Salt to apply when creating unique ID.
        private byte[] Salt = new byte[] {25, 139, 201, 13};

        private JsonResponse CallService(LocalCredentials credentials, object data)
        {
            // Calls the local service to get the response for the user.
            return null;
        }

        private JsonResponse GetCredentialsResponse()
        {
            // Creates a response that tells the Outlook add-in to
            // request the user's credentials for the service.
            return null;
        }

        private bool LocalCredentialsAreValid(LocalCredentials credentials)
        {
            // Returns true if the service recognizes the credentials provided.
            return false;
        }

        private string ComputeSHA256Hash(string uniqueId, string authenticationMetadataUrl, byte[] salt)
        {
            byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(uniqueId, authenticationMetadataUrl));

            // Combine input bytes and salt.
            byte[] saltedInput = new byte[salt.Length + inputBytes.Length];
            salt.CopyTo(saltedInput, 0);
            inputBytes.CopyTo(saltedInput, salt.Length);

            // Compute the unique key.
            byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

            // Convert the hashed value to a string and return.
            return BitConverter.ToString(hashedBytes);
        }

        public JsonResponse GetResponseFromService(IdentityToken token, LocalCredentials credentials, object data)
        {
            JsonResponse response = null;
            // This method should never be called with a null token.
            if (null == token)
            {
                throw new ArgumentNullException("token");
            }

            if (null == credentials)
            {
                string uniqueKey = ComputeSHA256Hash(token.ExchangeID, token.AuthenticationMetadataUrl, Salt);
                if (!AuthenticationCache.ContainsKey(uniqueKey))
                {
                    // The user's credentials are not in the authentication cache. Ask
                    // for the credentials.
                    response = GetCredentialsResponse();
                }
                else
                {
                    // The user's credentials are in the cache; make a request.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials. For example,
                        // the user has ended their subscription to the service, or the
                        // credentials have expired. Get new credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // The service returned a response to the user. Return the
                        // service response.
                        response = serviceResponse;
                    }
                }
            }
            else
            {
                // If the credentials are not null, it's a request to add an identity
                // to the authentication cache. Check to determine whether the local credentials
                // sent to the service are known.
                if (LocalCredentialsAreValid(credentials))
                {
                    // The local credentials are known. Add them to the 
                    // cached credentials.
                    string uniqueKey = ComputeSHA256Hash(token.ExchangeID, token.AuthenticationMetadataUrl, Salt);
                    AuthenticationCache.Add(uniqueKey, credentials);

                    // Get a response from the service.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // Return the service response to the user.
                        response = serviceResponse;
                    }
                }
            }

            return response;
        }
    }}
```


## Authentification d’un utilisateur auprès de la bibliothèque de validation gérée


Si vous utilisez la bibliothèque gérée pour valider les jetons d’identité, il n’est pas nécessaire de calculer la clé unique. La propriété  **UniqueUserIdentification** de la classe **AppIdentityToken** peut être utilisée directement en tant que clé unique pour l’utilisateur. L’exemple de code suivant montre les modifications à apporter à la méthode **GetResponseFromService** dans l’exemple précédent afin d’utiliser la classe **AppIdentityToken**.


```js
        public JsonResponse GetResponseFromService(AppIdentityToken token, LocalCredentials credentials, object data)
        {
            JsonResponse response = null;
            // This method should never be called with a null token.
            if (null == token)
            {
                throw new ArgumentNullException("token");
            }

            if (null == credentials)
            {
                string uniqueKey = token.UniqueUserIdentitification;
                if (!AuthenticationCache.ContainsKey(uniqueKey))
                {
                    // The user's credentials are not in the authentication cache. Ask
                    // for the credentials.
                    response = GetCredentialsResponse();
                }
                else
                {
                    // User's credentials are in the cache. Make a request.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials. For example,
                        // the user has ended their subscription to the service, or the
                        // credentials have expired. Get new credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // The service returned a response to the user. Return the
                        // service response.
                        response = serviceResponse;
                    }
                }
            }
            else
            {
                // If the credentials are not null, it's a request to add an identity
                // to the authentication cache. Check to determine whether the local credentials
                // sent to the service are known.
                if (LocalCredentialsAreValid(credentials))
                {
                    // The local credentials are known. Add them to the 
                    // cached credentials. 
                    string uniqueKey = token.UniqueUserIdentitification;
                    AuthenticationCache.Add(uniqueKey, credentials);

                    // Get a response from the service.
                    var serviceResponse = CallService(AuthenticationCache[uniqueKey], data);

                    if (null == serviceResponse)
                    {
                        // There was a problem with the stored credentials.
                        response = GetCredentialsResponse();
                    }
                    else
                    {
                        // Return the service response to the user.
                        response = serviceResponse;
                    }
                }
            }

            return response;
        }
```


## Ressources supplémentaires



- [Authentifier un complément Outlook à l’aide de jetons d’identité Exchange](../outlook/authentication.md)
    
- [Appeler un service à partir d’un complément Outlook à l’aide d’un jeton d’identité dans Exchange](../outlook/call-a-service-by-using-an-identity-token.md)
    
- [Utiliser la bibliothèque de validation des jetons Exchange](../outlook/use-the-token-validation-library.md)
    
- [Valider un jeton d’identité Exchange](../outlook/validate-an-identity-token.md)
    
