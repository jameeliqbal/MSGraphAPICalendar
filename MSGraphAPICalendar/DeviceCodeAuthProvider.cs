
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
 
namespace MSGraphAPICalendar.Graph.Authentication
{
    public class DeviceCodeAuthProvider : IAuthenticationProvider
    {
        private readonly IPublicClientApplication _msalClient;
        private readonly IEnumerable<string> _scopes;
        private IAccount _userAccount;

        public DeviceCodeAuthProvider(string appId, IEnumerable<string> scopes)
        {
            _scopes = scopes;

            _msalClient = PublicClientApplicationBuilder
                .Create(appId)
                .WithAuthority(AadAuthorityAudience
                    .AzureAdAndPersonalMicrosoftAccount, true)
                .Build();
        }

        public async Task<string> GetAccessToken()
        {
            // If there is no saved user account, the user must sign-in
            if (_userAccount == null)
            {
                try
                {
                    Console.WriteLine("SCOPES:..."+_scopes.ToString());
                    // Invoke device code flow so user can sign-in with a browser
                    var result = await _msalClient.AcquireTokenWithDeviceCode(
                        _scopes, callback => {
                            Console.WriteLine(callback.Message);
                            return Task.FromResult(0);
                        }).ExecuteAsync().ConfigureAwait(false);

                    _userAccount = result.Account;
                    return result.AccessToken;
                }
                catch (Exception exception)
                {
                    Console.WriteLine($"Error getting access token: {exception.Message}");
                    return $"Error getting access token: {exception.Message} {Environment.NewLine} {exception.InnerException?.Message}";
                }
            }
            else
            {
                // If there is an account, call AcquireTokenSilent
                // By doing this, MSAL will refresh the token automatically if
                // it is expired. Otherwise it returns the cached token.

                var result = await _msalClient
                    .AcquireTokenSilent(_scopes, _userAccount)
                    .ExecuteAsync();

                return result.AccessToken;
            }
        }

        // This is the required function to implement IAuthenticationProvider
        // The Graph SDK will call this function each time it makes a Graph
        // call.
        public async Task AuthenticateRequestAsync(HttpRequestMessage requestMessage)
        {
            requestMessage.Headers.Authorization =
                new AuthenticationHeaderValue("bearer", await GetAccessToken());
        }
    }
}