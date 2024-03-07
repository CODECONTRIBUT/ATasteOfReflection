using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ATasteOfReflection
{
    public class AttachmentModel
    {
        public int JobId { get; set; }

        public int AttachmentId { get; set; }

        public string AttachmentName { get; set; }

        public string AttachmentPath { get; set; }

        public AttachmentModel() 
        {
            using (var db = new DataContext())
            {
                if (AttachmentId == 0)
                    return;

                var attachment = db.Attachments.Where(m => m.AttachmentId = this.AttachmentId).FirstOrDefault();
                this.AttachmentName = attachment?.AttachmentName;
                this.AttachmentPath = attachment?.AttachmentPath;
            }
        }
    }
}
