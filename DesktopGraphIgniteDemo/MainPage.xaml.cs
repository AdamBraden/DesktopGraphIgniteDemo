using Microsoft.Graph;
using Microsoft.Toolkit.Services.MicrosoftGraph;
using Microsoft.Toolkit.Uwp.UI.Controls.Graph;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace DesktopGraphIgniteDemo
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        private string ClientId = "c9525f43-9dae-469d-b8fb-70f8e6fad6b0";
        private string[] permissions = new string[] { "User.Read", "Mail.Send" };
        public MainPage()
        {
            this.InitializeComponent();
            MicrosoftGraphService.Instance.AuthenticationModel = MicrosoftGraphEnums.AuthenticationModel.V2;
            MicrosoftGraphService.Instance.Initialize(
                ClientId,
                MicrosoftGraphEnums.ServicesToInitialize.UserProfile,
                permissions.Union(PeoplePicker.RequiredDelegatedPermissions).ToArray()
            );
            AadLogin1.View = ViewType.SmallProfilePhotoLeft;
            
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            var emailSubject = "Graph Email";

            var emailRecipients = new List<Microsoft.Graph.Recipient>();
            foreach (Microsoft.Graph.Person u in PeoplePicker1.Selections)
            {
                var recip = new Microsoft.Graph.Recipient
                {
                    EmailAddress = new EmailAddress() { Address = u.UserPrincipalName }
                };
                emailRecipients.Add(recip);
            }
            await SendMessageAsync(emailSubject, emailBody.Text, emailRecipients);
        }

        // Sends message to a specified address.
        public static async Task<bool> SendMessageAsync(string Subject, string Body, List<Recipient> Recipients)
        {
            bool emailSent = false;

            //Create message
            var email = new Message
            {
                Body = new ItemBody
                {
                    Content = Body,
                    ContentType = BodyType.Html,
                },
                Subject = Subject,
                ToRecipients = Recipients,
            };

            try
            {
                var graphClient = MicrosoftGraphService.Instance.GraphProvider;
                await graphClient.Me.SendMail(email, true).Request().PostAsync();
                Debug.WriteLine("Message sent");
                emailSent = true;
            }
            catch (ServiceException e)
            {
                Debug.WriteLine("We could not send the message. The request returned this status code: " + e.Error.Message);
                emailSent = false;
            }

            return emailSent;
        }

    }
}
