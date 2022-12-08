namespace Coginov.GraphApi.Library.Enums
{
    public enum AuthMethod
    {
        // Basic user and password
        Basic,

        //TODO: Research what OAuth flow is used here
        OAuthAppPermissions,
        
        //TODO: Research what OAuth flow is used here
        OAuthDelegatedPermissions,
        
        // This flow is a variant of OAuth 2.0 authorization code flow: https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow
        // We skip the authorization_code call to get a token and assume we already have valids JWT access and refresh tokens
        // This is a custom authentication method implemented without MSAL
        OAuthJwtAccessToken
    }
}