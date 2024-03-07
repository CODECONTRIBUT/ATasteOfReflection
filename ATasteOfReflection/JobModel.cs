using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ATasteOfReflection
{
    public class JobModel
    {
        public int JobId { get; set; }

        public string JobStatus { get; set; }

        public string JobManager { get; set; }

        public int JobAttachmentId { get; set; }

        private ProdDataContext _db;
        public ProdDataContext db
        {
            get
            {
                if (_db != null)
                    return _db;

                _db = new ProdDataContext();
                return _db;
            }
        }

        public JobModel()
        {
            //init
            if (JobId == 0)
                return;

            var job = db.Jobs.Where(m => m.Job = this.JobId).FirstOrDefault();
            this.JobStatus = job?.JobStatus;
            this.JobManager = job?.JobManager;
            this.JobAttachmentId = job?.JobAttachmentId;
        }
    }
}
