using System.Collections.Generic;
using System.Configuration;
using System.Linq;

namespace poc_outlook_mail_reader
{
    public static class ConfigHelper
    {
        public static string LogPath { get { return ConfigurationManager.AppSettings["LogPath"].ToString(); } }
        public static string EmailAddress { get { return ConfigurationManager.AppSettings["EmailAddress"].ToString(); } }
        public static string EmailPassword { get { return ConfigurationManager.AppSettings["EmailPassword"].ToString(); } }
        public static string ExchangeUrl { get { return ConfigurationManager.AppSettings["ExchangeUrl"].ToString(); } }
    }
}
