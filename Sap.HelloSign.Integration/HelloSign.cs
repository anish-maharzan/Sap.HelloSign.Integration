using HelloSign;
using Sap.HelloSign.Integration.Helper;
using System;
using System.IO;

namespace Sap.HelloSign.Integration
{
    public class HelloSign
    {
        private readonly Client client;
        public HelloSign()
        {
            client = new Client(AppConfiguration.API_KEY, AppConfiguration.API_HOST);
            //Test();
            //SendSignatureRequest();
            GetSignatureRequests();
        }

        public void Test()
        {
            // Using your account's API Key
            var client = new Client(AppConfiguration.API_KEY);
            var account = client.GetAccount();
            Console.WriteLine("My Account ID is: " + account.AccountId);

            account = client.UpdateAccount(new Uri("https://example.com/testhellosign.asp"));
            Console.WriteLine("Now my Callback URL is: " + account.CallbackUrl);
        }

        public void GetSignatureRequests()
        {
            // List signature requests
            var requests = client.ListSignatureRequests();
            Console.WriteLine("Found this many signature requests: " + requests.NumResults);
            foreach (var result in requests)
            {
                Console.WriteLine("Signature request: " + result.SignatureRequestId);
            }
        }

        public void SendSignatureRequest()
        {
            byte[] file1= File.ReadAllBytes("D:\\sapcco\\Files\\picklist.png");

            var request = new SignatureRequest();
            request.Title = "HelloSign demo";
            request.Subject = "HelloSign demo";
            request.Message = "Please sign in following documents";
            request.AddSigner("maharjananis1011@gmail.com", "Anish");
            request.AddSigner("anishmaharjan512@gmail.com", "Anish2");
            request.AddCc("anishmaharjan512@gmail.com");
            request.AddFile(file1, "picklist.png");
            request.Metadata.Add("custom_id", "1234");
            //client.AdditionalParameters.Add("metadata[custom_text]", "NDA #9"); // Inject additional parameter by hand
            request.AllowDecline = true;
            request.SigningOptions = new SigningOptions
            {
                Draw = true,
                Type = true,
                Default = "type"
            };
            request.TestMode = true;
            var response = client.SendSignatureRequest(request);
            Console.WriteLine("New Signature Request ID: " + response.SignatureRequestId);
        }
    }
}
