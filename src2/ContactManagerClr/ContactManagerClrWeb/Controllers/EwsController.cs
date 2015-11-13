using Microsoft.Exchange.WebServices.Auth.Validation;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web.Http;
using System.Web.Script.Serialization;

using Newtonsoft.Json;

namespace ContactManagerClrWeb.Controllers
{
    public class EwsController : ApiController
    {

        [HttpPost]
        public HttpResponseMessage CreateContact(CapturedEmailMessage capturedEmailMessage)
        {
            var response = Request.CreateResponse(HttpStatusCode.OK);

            try
            {
                // ensure we have a valid Identity Token
                var idToken = (AppIdentityToken)AuthToken.Parse(capturedEmailMessage.IdentityToken);
                var token = TokenDecoder.Decode(capturedEmailMessage.IdentityToken);

                // the identity token is valid 
                // placeholder

                var service = new ExchangeService(ExchangeVersion.Exchange2013)
                {
                    Url = new Uri(capturedEmailMessage.EWSUrl),
                    Credentials = new OAuthCredentials(capturedEmailMessage.EmailToken)
                };
            }
            catch (Exception ex)
            {
                response = Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex);
            }

            return response;

        }
    }

    public class CapturedEmailMessage
    {
        public string EmailToken { get; set; }
        public string IdentityToken { get; set; }
        public string ItemId { get; set; }
        public string EWSUrl { get; set; }
    }

    internal class TokenDecoder
    {
        public static Encoding TextEncoding = Encoding.UTF8;

        private static char Base64PadCharacter = '=';
        private static char Base64Character62 = '+';
        private static char Base64Character63 = '/';
        private static char Base64UrlCharacter62 = '-';
        private static char Base64UrlCharacter63 = '_';

        private static byte[] DecodeBytes(string arg)
        {
            if (String.IsNullOrEmpty(arg))
            {
                throw new ApplicationException("String to decode cannot be null or empty.");
            }

            StringBuilder s = new StringBuilder(arg);
            s.Replace(Base64UrlCharacter62, Base64Character62);
            s.Replace(Base64UrlCharacter63, Base64Character63);

            int pad = s.Length % 4;
            s.Append(Base64PadCharacter, (pad == 0) ? 0 : 4 - pad);

            return Convert.FromBase64String(s.ToString());
        }

        private static string Base64Decode(string arg)
        {
            return TextEncoding.GetString(DecodeBytes(arg));
        }

        public static JsonToken Decode(string rawToken)
        {
            string[] tokenParts = rawToken.Split('.');

            if (tokenParts.Length != 3)
            {
                throw new ApplicationException("Token must have three parts separated by '.' characters.");
            }

            string encodedHeader = tokenParts[0];
            string encodedPayload = tokenParts[1];
            string signature = tokenParts[2];

            string decodedHeader = Base64Decode(encodedHeader);
            string decodedPayload = Base64Decode(encodedPayload);

            JavaScriptSerializer serializer = new JavaScriptSerializer();

            Dictionary<string, string> header = serializer.Deserialize<Dictionary<string, string>>(decodedHeader);
            Dictionary<string, string> payload = serializer.Deserialize<Dictionary<string, string>>(decodedPayload);

            return new JsonToken (header, payload, signature);
        }
    }

    internal class JsonToken
    {
        public bool IsValid;
        public Dictionary<string, string> headerClaims;
        public Dictionary<string, string> payloadClaims;
        public string signature;
        public Dictionary<string, string> appContext;

        private void ValidateHeaderClaim(string key, string value)
        {
            if (!this.headerClaims.ContainsKey(key))
            {
                throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", key));
            }

            if (!value.Equals(this.headerClaims[key]))
            {
                throw new ApplicationException(String.Format("\"{0}\" claim must be \"{0}\".", key, value));
            }
        }

        private void ValidateHeader()
        {
            ValidateHeaderClaim("typ", "JWT");
            ValidateHeaderClaim("alg", "RS256");

            if (!this.headerClaims.ContainsKey("x5t"))
            {
                throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", "x5t"));
            }
        }
        private void ValidateLifetime()
        {
            if (!this.payloadClaims.ContainsKey("nbf"))
            {
                throw new ApplicationException(
                  String.Format("The \"{0}\" claim is missing from the token.", "nbf"));
            }

            if (!this.payloadClaims.ContainsKey("exp"))
            {
                throw new ApplicationException(
                  String.Format("The \"{0}\" claim is missing from the token.", "exp"));
            }

            DateTime unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

            TimeSpan padding = new TimeSpan(0, 5, 0);

            DateTime validFrom = unixEpoch.AddSeconds(int.Parse(this.payloadClaims["nbf"]));
            DateTime validTo = unixEpoch.AddSeconds(int.Parse(this.payloadClaims["exp"]));

            DateTime now = DateTime.UtcNow;

            if (now < (validFrom - padding))
            {
                throw new ApplicationException(String.Format("The token is not valid until {0}.", validFrom));
            }

            if (now > (validTo + padding))
            {
                throw new ApplicationException(String.Format("The token is not valid after {0}.", validFrom));
            }
        }
        private void ValidateMetadataLocation()
        {
            if (!this.appContext.ContainsKey("amurl"))
            {
                throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", "amurl"));
            }
        }



        private void ValidateAudience()
        {
            if (!this.payloadClaims.ContainsKey("aud"))
            {
                throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the application context.", "aud"));
            }

        }



        public JsonToken(Dictionary<string, string> header, Dictionary<string, string> payload, string signature)
        {

            // Assume that the token is invalid to start out.
            this.IsValid = false;

            // Set the private dictionaries that contain the claims.
            this.headerClaims = header;
            this.payloadClaims = payload;
            this.signature = signature;

            // If there is no "appctx" claim in the token, throw an ApplicationException.
            if (!this.payloadClaims.ContainsKey("appctx"))
            {
                throw new ApplicationException(String.Format("The {0} claim is not present.", "appctx"));
            }

            appContext = new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(payload["appctx"]);


            // Validate the header fields.
            this.ValidateHeader();

            // Determine whether the token is within its valid time.
            this.ValidateLifetime();

            // Validate that the token was sent to the correct URL.
            this.ValidateAudience();

            // Make sure that the appctx contains an authentication
            // metadata location.
            this.ValidateMetadataLocation();

            // If the token passes all the validation checks, we
            // can assume that it is valid.
            this.IsValid = true;
        }

    }
}
