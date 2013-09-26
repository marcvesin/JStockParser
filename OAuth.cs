using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Threading.Tasks;
using Google.GData.Client;
using Google.GData.Spreadsheets;
using System.IO;
using System.Net;



namespace GoogleStockParser
{
    class OAuth
    {

    //    Dictionary<string,string> _client_informations;
    //    OAuth2Parameters _parameter;
    //    string _scope;

    //    public OAuth()
    //    {
    //        Hashtable config = (Hashtable)ConfigurationManager.GetSection("Google_Data/OAuth_config");
    //        Hashtable scopes = (Hashtable)ConfigurationManager.GetSection("Google_Data/Scopes");

    //        _client_informations = new Dictionary<string,string>();

    //        foreach (string key in config.Keys)
    //            _client_informations.Add(key,(string)config[key]);

    //        foreach (string key in scopes.Keys)
    //            _scope += scopes[key] + " ";

    //        _scope.Substring(0, _scope.Length - 1);


    //        _parameter = Get_Parameter();
    //    }

    //    public OAuth2Parameters Get_Parameter()
    //    {
    //        /*
    //          // OAuth2Parameters holds all the parameters related to OAuth 2.0.
    //          OAuth2Parameters parameters = new OAuth2Parameters()
    //              { 
    //                  ClientId = _client_informations["Client_ID"],
    //                  ClientSecret = _client_informations["Client_secret"],
    //                  RedirectUri = _client_informations["Redirect_URI"]
    //              };


    //          ////////////////////////////////////////////////////////////////////////////
    //          // STEP 3: Get the Authorization URL
    //          ////////////////////////////////////////////////////////////////////////////

    //          // Set the scope for this particular service.
    //          parameters.Scope = _scope;

    //          // Get the authorization url.  The user of your application must visit
    //          // this url in order to authorize with Google.  If you are building a
    //          // browser-based application, you can redirect the user to the authorization
    //          // url.
    //          string authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(parameters);

    //          HttpWebRequest hwr = (HttpWebRequest)WebRequest.Create(authorizationUrl);
    //          hwr.Proxy = null;
    //          HttpWebResponse myHttpWebResponse = (HttpWebResponse)hwr.GetResponse();

    //          StreamWriter str = new StreamWriter("test.txt");
    //          str.WriteLine(authorizationUrl);
    //          str.Close();
    //          Console.WriteLine(authorizationUrl);

    //          Console.WriteLine("Please visit the URL above to authorize your OAuth "
    //            + "request token.  Once that is complete, type in your access code to "
    //            + "continue...");
    //          parameters.AccessCode = Console.ReadLine();

    //          ////////////////////////////////////////////////////////////////////////////
    //          // STEP 4: Get the Access Token
    //          ////////////////////////////////////////////////////////////////////////////

    //          // Once the user authorizes with Google, the request token can be exchanged
    //          // for a long-lived access token.  If you are building a browser-based
    //          // application, you should parse the incoming request token from the url and
    //          // set it in OAuthParameters before calling GetAccessToken().
    //          OAuthUtil.GetAccessToken(parameters);
    //          string accessToken = parameters.AccessToken;
    //          Console.WriteLine("OAuth Access Token: " + accessToken);


    //          return parameters;*/

    //}

    }
}
//////////////////////////////////////////////////////////////////////////////
//// STEP 5: Make an OAuth authorized request to Google
//////////////////////////////////////////////////////////////////////////////

//// Initialize the variables needed to make the request
//GOAuth2RequestFactory requestFactory =
//    new GOAuth2RequestFactory(null, "MySpreadsheetIntegration-v1", parameters);
//SpreadsheetsService service = new SpreadsheetsService("MySpreadsheetIntegration-v1");
//service.RequestFactory = requestFactory;

//// Make the request to Google
//// See other portions of this guide for code to put here...


//OAuth 2.0 scope information for the Google Spreadsheets API:
//https://spreadsheets.google.com/feeds
//Additionally, if an application needs to create spreadsheets, or otherwise manipulate their metadata, then the application must also request the Google Documents Lists API scope:
//https://docs.google.com/feeds
