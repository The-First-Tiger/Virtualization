namespace Virtualization.HyperV
{
    using System;
    using System.Management;

    internal class MethodResult
    {
        private const int Successful = 0;
        private const int Job = 4096;

        internal ManagementBaseObject MethodResultObject { get; set; }

        internal MethodResult(ManagementBaseObject methodResult)
        {
            this.MethodResultObject = methodResult;
        }

        internal bool WasSuccessful
        {
            get
            {
                return (UInt16)this.MethodResultObject["ReturnValue"] == Successful;
            }
        }

        internal bool IsJob
        {
            get
            {
                return (UInt16)this.MethodResultObject["ReturnValue"] == Job;
            }
        }

        internal Job GetCompletedJob()
        {
            var job = this.GetJob();

            while (job.CurrentJobState == JobStates.Starting
                   || job.CurrentJobState == JobStates.Running)
            {
                job = this.GetJob();
            }

            return job;
        }

        internal Job GetJob()
        {
            return new Job(new ManagementObject(this.MethodResultObject["Job"].ToString()));
        }

        internal static void HandleResult(ManagementBaseObject methodResultObject)
        {
            var methodResult = new MethodResult(methodResultObject);

            if (!methodResult.WasSuccessful && !methodResult.IsJob)
            {
                throw new InvalidOperationException("Could not execute!");
            }

            if (!methodResult.IsJob)
            {
                return;
            }

            var job = methodResult.GetCompletedJob();

            if (!job.WasSuccessful)
            {
                throw new InvalidOperationException("Job failed: " + job.ErrorDescription);
            }
        }
    }
}
