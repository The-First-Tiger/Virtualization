namespace Virtualization.HyperV
{
    using System;
    using System.Management;

    internal class Job
    {
        private readonly ManagementObject jobObject;        

        public Job(ManagementObject job)
        {
            this.jobObject = job;
        }

        public string ErrorDescription
        {
            get
            {
                return this.jobObject["ErrorDescription"].ToString();
            }            
        }

        public bool WasSuccessful
        {
            get
            {
                return this.CurrentJobState == JobStates.Completed;
            }            
        }

        public JobStates CurrentJobState
        {
            get
            {
                return (JobStates)(UInt16)this.jobObject["JobState"];
            }            
        }
    }
}
