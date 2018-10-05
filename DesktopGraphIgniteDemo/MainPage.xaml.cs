using System;
using System.Threading.Tasks;
using System.Linq;
using Windows.ApplicationModel.UserActivities;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Navigation;
using Microsoft.Graph;
using Microsoft.Toolkit.Services.MicrosoftGraph;

namespace DesktopGraphIgniteDemo
{
    /// <summary>
    /// Demo showing MSGraph, Windows Community Toolkit, and Windows Timeline
    /// </summary>
    public sealed partial class MainPage : Page
    {
        // Privates for Graph authentication
        private string ClientId = "c9525f43-9dae-469d-b8fb-70f8e6fad6b0";
        private string[] Permissions = new string[] { "User.Read", "User.ReadBasic.All", "People.Read", "Mail.Send"};

        // Private for UserActivities/Windows Timeline
        private UserActivitySession _currentSession;

        public MainPage()
        {
            this.InitializeComponent();
            MicrosoftGraphService.Instance.AuthenticationModel = MicrosoftGraphEnums.AuthenticationModel.V2;
            MicrosoftGraphService.Instance.Initialize(
                ClientId,
                MicrosoftGraphEnums.ServicesToInitialize.UserProfile | MicrosoftGraphEnums.ServicesToInitialize.Message,
                Permissions
            );
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            //SendEmail via Windows Community Toolkit's MSGraphService
            await MicrosoftGraphService.Instance.User.Message.SendEmailAsync(
                "Graph email via Windows Community Toolkit",
                emailBody.Text,
                BodyType.Html,
                PeoplePicker1.Selections.Select(x => x.UserPrincipalName).ToArray()
                );

            //Update status and show in WindowsTimeline
            statusBar.Text += " Success!";
            await CreateUserActivityAsync();
        }


        async Task CreateUserActivityAsync()
        {
            // Get channel and create activity.
            UserActivityChannel channel = UserActivityChannel.GetDefault();
            UserActivity activity = await channel.GetOrCreateUserActivityAsync("SentEmail");

            // Set deep-link and properties.
            activity.VisualElements.DisplayText = "Graph Demo - Email Sent!";
            activity.ActivationUri = new Uri("ignitedemo://page?MainPage");

            // Save to activity feed.
            await activity.SaveAsync();

            // Create a session, which indicates that the user is engaged in the activity.
            _currentSession?.Dispose();
            _currentSession = activity.CreateSession();
        }

        protected override void OnNavigatedFrom(NavigationEventArgs e)
        {
            // Dispose the session, which indicates that the user is no longer
            // engaged in the activity.
            _currentSession?.Dispose();
        }

        //// Sends message to a specified address.
        //public static async Task<bool> SendMessageAsync(string Subject, string Body, List<Recipient> Recipients)
        //{
        //    bool emailSent = false;

        //    //Create message
        //    var email = new Message
        //    {
        //        Body = new ItemBody {Content = Body, ContentType = BodyType.Html},
        //        Subject = Subject,
        //        ToRecipients = Recipients,
        //    };

        //    try
        //    {
        //        var graphClient = MicrosoftGraphService.Instance.GraphProvider;

        //        // Call the graph!
        //        await graphClient.Me.SendMail(email, true).Request().PostAsync();
        //        emailSent = true;
        //    }
        //    catch (ServiceException e)
        //    {
        //        Debug.WriteLine("We could not send the message. The request returned this status code: " + e.Error.Message);
        //        emailSent = false;
        //    }
        //    return emailSent;
        //}
    }
}
