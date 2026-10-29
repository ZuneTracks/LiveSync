using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using Windows.Data.Json;
using Windows.Security.Authentication.Web;
using Windows.Security.Credentials;
using Windows.Security.Cryptography;
using Windows.Security.Cryptography.Core;
using Windows.Storage.Streams;

namespace GraphMobileSync.Services
{
    internal static class OAuthService
    {
        private const string RedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient";
        private const string ReadScope =
            "https://graph.microsoft.com/Contacts.Read offline_access User.Read Calendars.Read";
        private const string ReadWriteScope =
            "https://graph.microsoft.com/Contacts.ReadWrite offline_access User.Read Calendars.ReadWrite";
        private const string RefreshTokenCredentialResource = "LiveSync.MicrosoftGraph";
        private const string RefreshTokenCredentialUserName = "RefreshToken";

        public static void ValidateSettings(string clientId, string tenant)
        {
            Guid ignored;

            if (!Guid.TryParse(clientId, out ignored))
            {
                throw new InvalidOperationException("Enter a valid Microsoft Entra application (client) ID.");
            }

            if (string.IsNullOrWhiteSpace(tenant))
            {
                throw new InvalidOperationException("Enter a tenant ID, tenant domain, or common.");
            }
        }

        public static async Task<string> SignInAsync(
            string clientId,
            string tenant,
            bool requestWriteAccess = false,
            bool promptForAccountSelection = false)
        {
            string codeVerifier = CreateCodeVerifier();
            string promptParameter = requestWriteAccess
                ? "&prompt=consent"
                : promptForAccountSelection ? "&prompt=select_account" : string.Empty;
            string authorizeUrl = string.Format(
                "https://login.microsoftonline.com/{0}/oauth2/v2.0/authorize?client_id={1}&response_type=code&redirect_uri={2}&response_mode=query&scope={3}&code_challenge={4}&code_challenge_method=S256{5}",
                Uri.EscapeDataString(tenant),
                Uri.EscapeDataString(clientId),
                Uri.EscapeDataString(RedirectUri),
                Uri.EscapeDataString(GetScope(requestWriteAccess)),
                Uri.EscapeDataString(CreateCodeChallenge(codeVerifier)),
                promptParameter);

            WebAuthenticationResult result = await WebAuthenticationBroker.AuthenticateAsync(
                WebAuthenticationOptions.None,
                new Uri(authorizeUrl),
                new Uri(RedirectUri));

            if (result.ResponseStatus == WebAuthenticationStatus.UserCancel)
            {
                throw new InvalidOperationException("Sign-in was cancelled.");
            }

            if (result.ResponseStatus != WebAuthenticationStatus.Success)
            {
                throw new InvalidOperationException("Microsoft sign-in failed: " + result.ResponseErrorDetail);
            }

            string authorizationCode = GetQueryParameter(result.ResponseData, "code");

            if (string.IsNullOrEmpty(authorizationCode))
            {
                string error = GetQueryParameter(result.ResponseData, "error_description");
                throw new InvalidOperationException("Microsoft sign-in did not return an authorization code. " + error);
            }

            GraphToken token = await ExchangeCodeAsync(
                clientId,
                tenant,
                authorizationCode,
                codeVerifier,
                requestWriteAccess);
            SaveRefreshToken(token.RefreshToken);
            return token.AccessToken;
        }

        public static async Task<string> GetBackgroundAccessTokenAsync(bool requestWriteAccess = false)
        {
            string clientId = AppConfiguration.ClientId;
            string tenant = AppConfiguration.Tenant;
            ValidateSettings(clientId, tenant);
            string refreshToken = GetRefreshToken();

            if (string.IsNullOrEmpty(refreshToken))
            {
                throw new InvalidOperationException("No refresh token is available. Open LiveSync and sign in again.");
            }

            using (HttpClient client = new HttpClient())
            {
                Dictionary<string, string> parameters = new Dictionary<string, string>
                {
                    { "client_id", clientId },
                    { "grant_type", "refresh_token" },
                    { "refresh_token", refreshToken },
                    { "scope", GetScope(requestWriteAccess) }
                };

                GraphToken token = await RequestTokenAsync(client, tenant, parameters);

                if (!string.IsNullOrEmpty(token.RefreshToken))
                {
                    SaveRefreshToken(token.RefreshToken);
                }

                return token.AccessToken;
            }
        }

        public static bool HasSavedSession()
        {
            string clientId = AppConfiguration.ClientId;
            string tenant = AppConfiguration.Tenant;

            return !string.IsNullOrWhiteSpace(clientId)
                && !string.IsNullOrWhiteSpace(tenant)
                && TryGetRefreshTokenCredential(new PasswordVault()) != null;
        }

        public static void SignOut()
        {
            PasswordVault vault = new PasswordVault();
            PasswordCredential credential = TryGetRefreshTokenCredential(vault);

            if (credential != null)
            {
                vault.Remove(credential);
            }
        }

        private static string GetScope(bool requestWriteAccess)
        {
            return requestWriteAccess ? ReadWriteScope : ReadScope;
        }

        private static async Task<GraphToken> ExchangeCodeAsync(
            string clientId,
            string tenant,
            string authorizationCode,
            string codeVerifier,
            bool requestWriteAccess)
        {
            using (HttpClient client = new HttpClient())
            {
                Dictionary<string, string> parameters = new Dictionary<string, string>
                {
                    { "client_id", clientId },
                    { "grant_type", "authorization_code" },
                    { "code", authorizationCode },
                    { "redirect_uri", RedirectUri },
                    { "code_verifier", codeVerifier },
                    { "scope", GetScope(requestWriteAccess) }
                };

                return await RequestTokenAsync(client, tenant, parameters);
            }
        }

        private static async Task<GraphToken> RequestTokenAsync(
            HttpClient client,
            string tenant,
            Dictionary<string, string> parameters)
        {
            HttpResponseMessage response = await client.PostAsync(
                new Uri("https://login.microsoftonline.com/" + Uri.EscapeDataString(tenant) + "/oauth2/v2.0/token"),
                new FormUrlEncodedContent(parameters));
            string payload = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException("Token request failed: " + payload);
            }

            JsonObject token = JsonObject.Parse(payload);
            string accessToken = JsonValueReader.GetString(token, "access_token", string.Empty);

            if (string.IsNullOrEmpty(accessToken))
            {
                throw new InvalidOperationException("Microsoft Entra did not return an access token.");
            }

            return new GraphToken(
                accessToken,
                JsonValueReader.GetString(token, "refresh_token", string.Empty));
        }

        private static void SaveRefreshToken(string refreshToken)
        {
            if (string.IsNullOrEmpty(refreshToken))
            {
                throw new InvalidOperationException(
                    "Microsoft Entra did not return a refresh token. Confirm offline_access is permitted, then sign in again.");
            }

            PasswordVault vault = new PasswordVault();
            PasswordCredential existingCredential = TryGetRefreshTokenCredential(vault);

            if (existingCredential != null)
            {
                vault.Remove(existingCredential);
            }

            vault.Add(new PasswordCredential(
                RefreshTokenCredentialResource,
                RefreshTokenCredentialUserName,
                refreshToken));
        }

        private static string GetRefreshToken()
        {
            PasswordCredential credential = TryGetRefreshTokenCredential(new PasswordVault());

            if (credential == null)
            {
                return string.Empty;
            }

            credential.RetrievePassword();
            return credential.Password;
        }

        private static PasswordCredential TryGetRefreshTokenCredential(PasswordVault vault)
        {
            try
            {
                return vault.Retrieve(RefreshTokenCredentialResource, RefreshTokenCredentialUserName);
            }
            catch (Exception exception) when (exception.HResult == unchecked((int)0x80070490))
            {
                return null;
            }
        }

        private static string GetQueryParameter(string responseData, string name)
        {
            Uri responseUri = new Uri(responseData);
            string query = responseUri.Query.TrimStart('?');

            foreach (string pair in query.Split('&'))
            {
                string[] values = pair.Split(new[] { '=' }, 2);

                if (values.Length == 2 && values[0] == name)
                {
                    return Uri.UnescapeDataString(values[1].Replace("+", " "));
                }
            }

            return null;
        }

        private static string CreateCodeVerifier()
        {
            return Base64UrlEncode(CryptographicBuffer.GenerateRandom(32));
        }

        private static string CreateCodeChallenge(string codeVerifier)
        {
            HashAlgorithmProvider provider = HashAlgorithmProvider.OpenAlgorithm(HashAlgorithmNames.Sha256);
            IBuffer verifierBytes = CryptographicBuffer.ConvertStringToBinary(
                codeVerifier,
                BinaryStringEncoding.Utf8);
            return Base64UrlEncode(provider.HashData(verifierBytes));
        }

        private static string Base64UrlEncode(IBuffer bytes)
        {
            return CryptographicBuffer.EncodeToBase64String(bytes)
                .TrimEnd('=')
                .Replace('+', '-')
                .Replace('/', '_');
        }

        private sealed class GraphToken
        {
            public GraphToken(string accessToken, string refreshToken)
            {
                AccessToken = accessToken;
                RefreshToken = refreshToken;
            }

            public string AccessToken { get; private set; }

            public string RefreshToken { get; private set; }
        }
    }
}
