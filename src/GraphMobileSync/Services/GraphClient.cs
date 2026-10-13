using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Windows.Data.Json;

namespace GraphMobileSync.Services
{
    internal sealed class GraphDeltaResult
    {
        public GraphDeltaResult(IList<JsonObject> items, string deltaLink)
        {
            Items = items;
            DeltaLink = deltaLink;
        }

        public IList<JsonObject> Items { get; private set; }

        public string DeltaLink { get; private set; }
    }

    internal sealed class GraphClient
        : IDisposable
    {
        private const string OutlookTimeZonePreference = "outlook.timezone=\"UTC\"";
        private readonly HttpClient _httpClient;

        public GraphClient(string accessToken)
        {
            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentException("An access token is required.", nameof(accessToken));
            }

            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", accessToken);
            _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Prefer", OutlookTimeZonePreference);
        }

        public async Task<IList<JsonObject>> GetCollectionAsync(string relativeUrl)
        {
            List<JsonObject> items = new List<JsonObject>();
            Uri requestUri = CreateRequestUri(relativeUrl);

            while (requestUri != null)
            {
                JsonObject response = await GetAsync(requestUri);
                JsonArray values = JsonValueReader.GetArray(response, "value");

                foreach (IJsonValue value in values)
                {
                    if (value.ValueType == JsonValueType.Object)
                    {
                        items.Add(value.GetObject());
                    }
                }

                string nextLink = JsonValueReader.GetString(response, "@odata.nextLink", string.Empty);
                requestUri = string.IsNullOrEmpty(nextLink) ? null : new Uri(nextLink);
            }

            return items;
        }

        public async Task<GraphDeltaResult> GetDeltaAsync(string relativeOrAbsoluteUrl)
        {
            List<JsonObject> items = new List<JsonObject>();
            Uri requestUri = CreateRequestUri(relativeOrAbsoluteUrl);
            string deltaLink = string.Empty;

            while (requestUri != null)
            {
                JsonObject response = await GetAsync(requestUri);
                JsonArray values = JsonValueReader.GetArray(response, "value");

                foreach (IJsonValue value in values)
                {
                    if (value.ValueType == JsonValueType.Object)
                    {
                        items.Add(value.GetObject());
                    }
                }

                string nextLink = JsonValueReader.GetString(response, "@odata.nextLink", string.Empty);
                deltaLink = JsonValueReader.GetString(response, "@odata.deltaLink", deltaLink);
                requestUri = string.IsNullOrEmpty(nextLink) ? null : new Uri(nextLink);
            }

            if (string.IsNullOrEmpty(deltaLink))
            {
                throw new InvalidOperationException("Microsoft Graph did not return a delta cursor.");
            }

            return new GraphDeltaResult(items, deltaLink);
        }

        public async Task<byte[]> GetContactPhotoAsync(string contactId)
        {
            if (string.IsNullOrWhiteSpace(contactId))
            {
                throw new ArgumentException("A contact ID is required.", nameof(contactId));
            }

            Uri requestUri = CreateRequestUri(string.Format(
                "/v1.0/me/contacts/{0}/photo/$value",
                Uri.EscapeDataString(contactId)));
            TimeSpan deferredDelay;

            if (GraphThrottleState.TryGetRemainingDelay(out deferredDelay))
            {
                throw new GraphThrottledException(deferredDelay);
            }

            using (HttpResponseMessage response = await _httpClient.GetAsync(requestUri))
            {
                if ((int)response.StatusCode == 429)
                {
                    TimeSpan retryAfter = GetRetryAfter(response.Headers);
                    GraphThrottleState.Defer(retryAfter);
                    throw new GraphThrottledException(retryAfter);
                }

                if ((int)response.StatusCode == 404)
                {
                    return null;
                }

                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException(
                        "Microsoft Graph contact photo request failed: " + response.ReasonPhrase);
                }

                return await response.Content.ReadAsByteArrayAsync();
            }
        }

        public void Dispose()
        {
            _httpClient.Dispose();
        }

        private static Uri CreateRequestUri(string relativeOrAbsoluteUrl)
        {
            Uri absoluteUri;

            if (Uri.TryCreate(relativeOrAbsoluteUrl, UriKind.Absolute, out absoluteUri))
            {
                return absoluteUri;
            }

            return new Uri("https://graph.microsoft.com" + relativeOrAbsoluteUrl);
        }

        private async Task<JsonObject> GetAsync(Uri requestUri)
        {
            TimeSpan deferredDelay;

            if (GraphThrottleState.TryGetRemainingDelay(out deferredDelay))
            {
                throw new GraphThrottledException(deferredDelay);
            }

            using (HttpResponseMessage response = await _httpClient.GetAsync(requestUri))
            {
                if ((int)response.StatusCode == 429)
                {
                    TimeSpan retryAfter = GetRetryAfter(response.Headers);
                    GraphThrottleState.Defer(retryAfter);
                    throw new GraphThrottledException(retryAfter);
                }

                string payload = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException("Microsoft Graph request failed: " + payload);
                }

                return JsonObject.Parse(payload);
            }
        }

        private static TimeSpan GetRetryAfter(HttpResponseHeaders headers)
        {
            RetryConditionHeaderValue retryAfter = headers.RetryAfter;

            if (retryAfter != null)
            {
                if (retryAfter.Delta.HasValue)
                {
                    return retryAfter.Delta.Value;
                }

                if (retryAfter.Date.HasValue)
                {
                    TimeSpan remaining = retryAfter.Date.Value - DateTimeOffset.UtcNow;

                    if (remaining > TimeSpan.Zero)
                    {
                        return remaining;
                    }
                }
            }

            return TimeSpan.FromSeconds(60);
        }
    }
}
