//Author: Sergiy Stoyan
//        s.y.stoyan@gmail.com, sergiy.stoyan@outlook.com, stoyan@cliversoft.com
//        http://www.cliversoft.com
//********************************************************************************************
using System;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using Azure.Identity;
using Microsoft.Kiota.Abstractions.Authentication;
using Microsoft.Kiota.Abstractions;
using System.Threading;
using Microsoft.Kiota.Abstractions.Serialization;
using Microsoft.Kiota.Abstractions.Store;

namespace Cliver
{
    public class MicrosoftService
    {
        public MicrosoftService(MicrosoftSettings microsoftSettings)
        {
            MicrosoftSettings = microsoftSettings;
            Client = createClient();
        }
        public readonly MicrosoftSettings MicrosoftSettings;

        public string MicrosoftAccount
        {
            get
            {
                if (account == null)
                    Authenticate();
                return account?.Username;
            }
        }

        public GraphServiceClient Client { get; private set; }

        GraphServiceClient createClient()
        {
            application = PublicClientApplicationBuilder.Create(MicrosoftSettings.ClientId)
            .WithTenantId(MicrosoftSettings.TenantId)
            .WithRedirectUri("http://localhost")//to use the default browser
            .Build();

            application.UserTokenCache.SetAfterAccess(MicrosoftSettings.AfterAccessNotification);
            application.UserTokenCache.SetBeforeAccess(MicrosoftSettings.BeforeAccessNotification);
            //application.UserTokenCache.SetBeforeWrite((TokenCacheNotificationArgs a) => { });
            //application.UserTokenCache.SetCacheOptions(new CacheOptions { UseSharedCache = false });

            if (string.IsNullOrWhiteSpace(MicrosoftSettings.MicrosoftAccount))
                account = Task.Run(() => application.GetAccountsAsync()).Result.FirstOrDefault();
            else
                account = Task.Run(() => application.GetAccountsAsync()).Result.FirstOrDefault(a => a.Username == MicrosoftSettings.MicrosoftAccount);

            return new GraphServiceClient(httpClient, new AuthenticationProvider(this));
        }
        IPublicClientApplication application;
        IAccount account = null;
        System.Net.Http.HttpClient httpClient = GraphClientFactory.Create();
        public class AuthenticationProvider : IAuthenticationProvider
        {
            internal AuthenticationProvider(MicrosoftService microsoftService)
            {
                this.microsoftService = microsoftService;
            }
            MicrosoftService microsoftService;

            async Task IAuthenticationProvider.AuthenticateRequestAsync(RequestInformation request, Dictionary<string, object> additionalAuthenticationContext, CancellationToken cancellationToken)
            {
                await microsoftService.authenticate();
                request.Headers.Add("Authorization", "bearer " + microsoftService.authenticationResult.AccessToken);
            }
        }
        async Task authenticate()
        {
            try
            {
                authenticationResult = await application.AcquireTokenSilent(MicrosoftSettings.Scopes, account).ExecuteAsync();
            }
            catch (MsalUiRequiredException e)
            {
                //if (e.ErrorCode != MsalError.InvalidGrantError && e.ErrorCode != MsalError.UserNullError /* || e.Classification == UiRequiredExceptionClassification.None*/)
                //    throw;
                OnInteractiveAuthentication?.Invoke();
                authenticationResult = await application.AcquireTokenInteractive(MicrosoftSettings.Scopes)
                    //.WithUseEmbeddedWebView(true)!!!intermittently gives the error (even when running in an STA thread): ActiveX control '8856f961-340a-11d0-a96b-00c04fd705a2' cannot be instantiated because the current thread is not in a single-threaded apartment. 
                    .WithUseEmbeddedWebView(false)
                    .ExecuteAsync();
                account = authenticationResult?.Account;

                if (MicrosoftSettings.MicrosoftAccount != account.Username)
                {
                    MicrosoftSettings.MicrosoftAccount = account.Username;
                    MicrosoftSettings.Save();
                }
            }
        }
        AuthenticationResult authenticationResult = null;

        public Action OnInteractiveAuthentication = null;

        /*public void Authenticate2()//WithUseEmbeddedWebView(true)
        {
            Task.Run(async () => { await authenticate(); }).Wait();//!!!on the client's computer it gave: ActiveX control '8856f961-340a-11d0-a96b-00c04fd705a2' cannot be instantiated because the current thread is not in a single-threaded apartment. 

            //ThreadRoutines.StartTrySta(authenticate().Wait).Join();//!!!intermittently freezes

            //if (System.Threading.Thread.CurrentThread.GetApartmentState() == System.Threading.ApartmentState.STA)
            //TaskRoutines.RunSynchronously(authenticate);//!!!on the client's computer it gave: ActiveX control '8856f961-340a-11d0-a96b-00c04fd705a2' cannot be instantiated because the current thread is not in a single-threaded apartment. 
            //else
            //   ThreadRoutines.StartTrySta(() => { TaskRoutines.RunSynchronously(authenticate); }).Join();//feezes
        }*/
        public void Authenticate()//WithUseEmbeddedWebView(false)
        {
            //authenticate().Wait();//never returns from AcquireTokenInteractive()
            //Task.Run(async () => { await authenticate(); }).Wait();//freezes at OnInteractiveAuthentication() 
            //Task.Run(() => { authenticate(); }).Wait();//freezes at OnInteractiveAuthentication() 

            //ThreadRoutines.StartTrySta(authenticate().Wait).Join();//!!!intermittently freezes

            //if (System.Threading.Thread.CurrentThread.GetApartmentState() == System.Threading.ApartmentState.STA)
            TaskRoutines.RunSynchronously(authenticate);//???works reliably?
            //else
            //   ThreadRoutines.StartTrySta(() => { TaskRoutines.RunSynchronously(authenticate); }).Join();//feezes
        }

        public TimeSpan Timeout
        {
            get
            {
                return httpClient.Timeout;
            }
            set
            {
                httpClient.Timeout = value;
            }
        }

        public User GetUser(string userId = null)
        {
            return Task.Run(() =>
            {
                if (userId == null)
                    return Client.Me.GetAsync();
                else
                    return Client.Users[userId].GetAsync();
            }).Result;
        }

        public User User
        {
            get
            {
                if (user == null)
                    user = GetUser(null);
                return user;
            }
        }
        User user = null;
    }
}