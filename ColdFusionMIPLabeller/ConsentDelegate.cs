using Microsoft.InformationProtection;

namespace ColdFusionMIPLabeller
{
    /// <summary>
    /// Consent delegate for headless operation - always accepts consent requests.
    /// </summary>
    public class ConsentDelegate : IConsentDelegate
    {
        public Consent GetUserConsent(string url)
        {
            // Headless operation - always accept consent
            return Consent.AcceptAlways;
        }
    }
}