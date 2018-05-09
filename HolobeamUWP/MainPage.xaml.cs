using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using Microsoft.Identity.Client;
using System.Threading.Tasks;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace HolobeamUWP
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        //Set the API Endpoint to Graph 'me' endpoint
        string graphAPIEndpoint = "https://graph.microsoft.com/v1.0/me";

        //Set the scope for API call to user.read
        string[] scopes = new string[] { "user.read" };

        public MainPage()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// Call AcquireTokenAsync - to acquire a token requiring user to sign-in
        /// </summary>
        private async void CallGraphButton_Click(object sender, RoutedEventArgs e)
        {
            AuthenticationResult authResult = null;
            ResultText.Text = string.Empty;
            TokenInfoText.Text = string.Empty;

            try
            {
                authResult = await App.PublicClientApp.AcquireTokenSilentAsync(scopes, App.PublicClientApp.Users.FirstOrDefault());
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilentAsync. This indicates you need to call AcquireTokenAsync to acquire a token
                System.Diagnostics.Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");   

                try
                {
                    authResult = await App.PublicClientApp.AcquireTokenAsync(scopes);
                }
                catch (MsalException msalex)
                {
                    ResultText.Text = $"Error Acquiring Token:{System.Environment.NewLine}{msalex}";
                }
            }
            catch (Exception ex)
            {
                ResultText.Text = $"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}";
                return;
            }

            if (authResult != null)
            {
                string Accesstoken = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFEWDhHQ2k2SnM2U0s4MlRzRDJQYjdyLXRTd09kd1hlSTJIV2JaWDdPUm9uQkhKQmgwR0ZoWE9QcGdENWRONWFrYk54YUhOUHlvV0NsVjdhUmFUYWNfNmZiRW5KRm0tNVVSSEpUb1JVMllsa0NBQSIsImFsZyI6IlJTMjU2IiwieDV0IjoiaUJqTDFSY3F6aGl5NGZweEl4ZFpxb2hNMllrIiwia2lkIjoiaUJqTDFSY3F6aGl5NGZweEl4ZFpxb2hNMllrIn0.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC81YzgwODVkOS0xZTg4LTRiYjYtYjViZC1lNmU2ZDViNWJhYmQvIiwiaWF0IjoxNTI1MzQ3NzI1LCJuYmYiOjE1MjUzNDc3MjUsImV4cCI6MTUyNTM1MTYyNSwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhIQUFBQWRYcVJRWUhyU0VKSmpIa2JUVDNNUWN6eHNYeStlMkg5MVZDY0ZESk1hUHhvRDdPNFk1b2F1NXdjc0FWeDhhUXlEVWduWHY2K2FqUWhhWExPN1VzQnVSL2ViVm5MUG05L1FUWVVJNEVPNVpZPSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiU2FtcGxlQXV0aEFwcCIsImFwcGlkIjoiZWFlOTdkMjEtYTk2OS00MThiLTgxNzUtOWI0YTk3NzgwZDEyIiwiYXBwaWRhY3IiOiIxIiwiZV9leHAiOjI2MjgwMCwiZmFtaWx5X25hbWUiOiJLIEoiLCJnaXZlbl9uYW1lIjoiQWtoaWwiLCJpcGFkZHIiOiI0OS4yNDkuMjMxLjE0NiIsIm5hbWUiOiJBa2hpbCBLIEoiLCJvaWQiOiJmMDczOGZhZi04NjE4LTQ2MzItOWFiZC1lMzc3MGRkMDI4ZWEiLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtMjg3MzM0NTk4OS0yMTI3NDgwNjIxLTQ2NjI5MTk5LTY4NDEiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzNGRkZBNzJFNEVDRiIsInNjcCI6Ik1haWwuUmVhZCBNYWlsLlNlbmQgVXNlci5SZWFkIiwic3ViIjoiOC03WGx1RVhoOTJaZlhMWUpXU195T2w4dEtkV0lyMXJ3UXpSSVlSa0FudyIsInRpZCI6IjVjODA4NWQ5LTFlODgtNGJiNi1iNWJkLWU2ZTZkNWI1YmFiZCIsInVuaXF1ZV9uYW1lIjoiYWtoaWxrakB2YWxvcmVtLmNvbSIsInVwbiI6ImFraGlsa2pAdmFsb3JlbS5jb20iLCJ1dGkiOiJrQkdJWVJFVDlVR2FSNXluZjRNSUFBIiwidmVyIjoiMS4wIn0.nA47tK-SemyFJcL8ZgaWgydlN3NIOnI9OkKzB19wKeCAHu2dElHNOQVvAQO3ytVGCmaBqmyUk88BEWfFtdHgAV2jx589o__Eb0SsRbxQ0lPiZp1esyJNIl5w5xM89goGpx5LLCDGsSxCbyYnaN472E2Gxw5tt9EXYA4uPtU4_61t4PiKvTcl9d1YLpLFxJX-cl9aAeZPiDHDp6Ok8y2DZ-6JRPnMBaMiHLmCtup7OLNLVnQUNGR_3PUQoaG1hmBWkbaZze5aRcjl7TMAR-KBv_5JyLwhLLOK1DtP2kuulREY0YBwU-n-2gsRxKjooR6n8WA5TjxpHYPuobmaBz2mQA";
                ResultText.Text = await GetHttpContentWithToken(graphAPIEndpoint, Accesstoken);
                DisplayBasicTokenInfo(authResult);
                this.SignOutButton.Visibility = Visibility.Visible;
            }
        }
        /// <summary>
        /// Perform an HTTP GET request to a URL using an HTTP Authorization header
        /// </summary>
        /// <param name="url">The URL</param>
        /// <param name="token">The token</param>
        /// <returns>String containing the results of the GET operation</returns>
        public async Task<string> GetHttpContentWithToken(string url, string token)
        {
            var httpClient = new System.Net.Http.HttpClient();
            System.Net.Http.HttpResponseMessage response;
            //token = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFEWDhHQ2k2SnM2U0s4MlRzRDJQYjdyLXRTd09kd1hlSTJIV2JaWDdPUm9uQkhKQmgwR0ZoWE9QcGdENWRONWFrYk54YUhOUHlvV0NsVjdhUmFUYWNfNmZiRW5KRm0tNVVSSEpUb1JVMllsa0NBQSIsImFsZyI6IlJTMjU2IiwieDV0IjoiaUJqTDFSY3F6aGl5NGZweEl4ZFpxb2hNMllrIiwia2lkIjoiaUJqTDFSY3F6aGl5NGZweEl4ZFpxb2hNMllrIn0.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC81YzgwODVkOS0xZTg4LTRiYjYtYjViZC1lNmU2ZDViNWJhYmQvIiwiaWF0IjoxNTI1MzQ3NzI1LCJuYmYiOjE1MjUzNDc3MjUsImV4cCI6MTUyNTM1MTYyNSwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhIQUFBQWRYcVJRWUhyU0VKSmpIa2JUVDNNUWN6eHNYeStlMkg5MVZDY0ZESk1hUHhvRDdPNFk1b2F1NXdjc0FWeDhhUXlEVWduWHY2K2FqUWhhWExPN1VzQnVSL2ViVm5MUG05L1FUWVVJNEVPNVpZPSIsImFtciI6WyJwd2QiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiU2FtcGxlQXV0aEFwcCIsImFwcGlkIjoiZWFlOTdkMjEtYTk2OS00MThiLTgxNzUtOWI0YTk3NzgwZDEyIiwiYXBwaWRhY3IiOiIxIiwiZV9leHAiOjI2MjgwMCwiZmFtaWx5X25hbWUiOiJLIEoiLCJnaXZlbl9uYW1lIjoiQWtoaWwiLCJpcGFkZHIiOiI0OS4yNDkuMjMxLjE0NiIsIm5hbWUiOiJBa2hpbCBLIEoiLCJvaWQiOiJmMDczOGZhZi04NjE4LTQ2MzItOWFiZC1lMzc3MGRkMDI4ZWEiLCJvbnByZW1fc2lkIjoiUy0xLTUtMjEtMjg3MzM0NTk4OS0yMTI3NDgwNjIxLTQ2NjI5MTk5LTY4NDEiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzNGRkZBNzJFNEVDRiIsInNjcCI6Ik1haWwuUmVhZCBNYWlsLlNlbmQgVXNlci5SZWFkIiwic3ViIjoiOC03WGx1RVhoOTJaZlhMWUpXU195T2w4dEtkV0lyMXJ3UXpSSVlSa0FudyIsInRpZCI6IjVjODA4NWQ5LTFlODgtNGJiNi1iNWJkLWU2ZTZkNWI1YmFiZCIsInVuaXF1ZV9uYW1lIjoiYWtoaWxrakB2YWxvcmVtLmNvbSIsInVwbiI6ImFraGlsa2pAdmFsb3JlbS5jb20iLCJ1dGkiOiJrQkdJWVJFVDlVR2FSNXluZjRNSUFBIiwidmVyIjoiMS4wIn0.nA47tK-SemyFJcL8ZgaWgydlN3NIOnI9OkKzB19wKeCAHu2dElHNOQVvAQO3ytVGCmaBqmyUk88BEWfFtdHgAV2jx589o__Eb0SsRbxQ0lPiZp1esyJNIl5w5xM89goGpx5LLCDGsSxCbyYnaN472E2Gxw5tt9EXYA4uPtU4_61t4PiKvTcl9d1YLpLFxJX-cl9aAeZPiDHDp6Ok8y2DZ-6JRPnMBaMiHLmCtup7OLNLVnQUNGR_3PUQoaG1hmBWkbaZze5aRcjl7TMAR-KBv_5JyLwhLLOK1DtP2kuulREY0YBwU-n-2gsRxKjooR6n8WA5TjxpHYPuobmaBz2mQA";
            try
            {
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url);
                //Add the token in Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                response = await httpClient.SendAsync(request);
                var content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }
        /// <summary>
        /// Sign out the current user
        /// </summary>
        private void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            if (App.PublicClientApp.Users.Any())
            {
                try
                {
                    App.PublicClientApp.Remove(App.PublicClientApp.Users.FirstOrDefault());
                    this.ResultText.Text = "User has signed-out";
                    this.CallGraphButton.Visibility = Visibility.Visible;
                    this.SignOutButton.Visibility = Visibility.Collapsed;
                }
                catch (MsalException ex)
                {
                    ResultText.Text = $"Error signing-out user: {ex.Message}";
                }
            }
        }
        /// <summary>
        /// Display basic information contained in the token
        /// </summary>
        private void DisplayBasicTokenInfo(AuthenticationResult authResult)
        {
            TokenInfoText.Text = "";
            if (authResult != null)
            {
                TokenInfoText.Text += $"Name: {authResult.User.Name}" + Environment.NewLine;
                TokenInfoText.Text += $"Username: {authResult.User.DisplayableId}" + Environment.NewLine;
                TokenInfoText.Text += $"Token Expires: {authResult.ExpiresOn.ToLocalTime()}" + Environment.NewLine;
                TokenInfoText.Text += $"Access Token: {authResult.AccessToken}" + Environment.NewLine;
            }
        }
    }
}
