using Microsoft.Azure.WebJobs;

namespace Bugfree.Spo.ExternalSharingCenter.WebJob
{
    class Program
    {
        static void Main(string[] args)
        {
            var host = new JobHost();
            host.Call(typeof(ExternalSharing).GetMethod("Start"));
        }
    }
}
