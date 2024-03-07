using System;
using System.Collections.Generic;
using System.IO;
using ATasteOfReflection;
using System.Reflection;

//We do not add Syncfusion references here as they are not free. Just put needed code here.
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Xml.Linq;


public class ReportModel
{
    public int JobId { get; set; }
    //instance Variables

    //Here WordDocument type and other doc types are from Syncfusion Word Library
    private WordDocument _reportDoc = null;
    public WordDocument ReportDoc 
    {
        get
        {
            if (_reportDoc != null)
                return _reportDoc;

            //In production env, usually get template path from database with template id,
            //but here we just focus on C# Reflection. So just hard code template path instead.
            var fileFullPath = @"C://Users/Template.docx";
            FileStream fileStreamPath = new FileStream(fileFullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            _reportDoc = new WordDocument(fileStreamPath, FormatType.Docx);
            return _reportDoc;
        }   
     }

    private JobModel _jobModel;
    public ATasteOfReflection.JobModel JobModel
    { 
        get
        {
            if (_jobModel != null)
                return _jobModel;

            if (this.JobId != 0)
            {
                var _jobModel = new JobModel() { JobId = this.JobId };
                return _jobModel;
            }
            else
                return null;
        }
    }



    public byte[] GenerateReportWithReflection()
    {
        try
        {
            //Get Tags for replacement in template
            var tags = GetAllTagesFromTemplate();

            //Replace tags with C# Reflection
            GenerateReportAndReplaceTags(tags);

            //Save
            return SaveReport();
        }
        catch (Exception ex)
        {
            Elmah.ErrorSignal.FromCurrentContext().Raise(ex);
            return null;
        }
    }

	private List<string> GetAllTagesFromTemplate()
	{
		try
		{
            var Tags = new List<string>();

            //Finds all the occurrences of Regex Expression, like "[JobId]"
            var regexEx = @"\[(.*?)\]";
            TextSelection[] textSelections = ReportDoc.FindAll(regexEx, false, true);
                
            foreach (var textSelection in textSelections)
            {
                Tags.Add(textSelection.ToString());
            }

            return Tags;
        }
		catch (Exception ex)
        {
            Elmah.ErrorSignal.FromCurrentContext().Raise(ex);
            return null;
        }
	}

    private void GenerateReportAndReplaceTags(List<string> tags)
    {
        try
        {
            var tagDic = InitializeTagDictionary();
            if (tagDic == null)
                return;

            var model = new ReportModel();
            foreach (var tag in tags)
            {
                var methodName = string.Empty;
                if (tagDic.ContainsKey(tag))
                {
                    methodName = tagDic[tag];
                    InvokeMethod(model, methodName, tag);
                }
            }
        }
        catch (Exception ex)
        {
            Elmah.ErrorSignal.FromCurrentContext().Raise(ex);
        }
    }

    private Dictionary<string, string> InitializeTagDictionary()
    {
        try
        {
            var tagDictonary = new Dictionary<string, string>();
            tagDictonary.Add("[JobId]", "ReplaceTagsWithJobModel");
            tagDictonary.Add("[JobStatus]", "ReplaceTagsWithJobModel");
            tagDictonary.Add("[JobManager]", "ReplaceTagsWithJobModel");
            tagDictonary.Add("[AttachmentName]", "ReplaceTagsWithAttachmentModel");
            tagDictonary.Add("[AttachmentPath]", "ReplaceTagsWithAttachmentModel");
            return tagDictonary;
        }
        catch (Exception ex)
        {
            Elmah.ErrorSignal.FromCurrentContext().Raise(ex);
            return null;
        }
    }

    private void InvokeMethod(ReportModel model, string methodName, string tag)
    {
        try
        {
            if (string.IsNullOrEmpty(methodName))
                return;

            var reportModelType = Type.GetType("ATasteOfReflection.ReportModel");
            var method = reportModelType.GetMethod(methodName, BindingFlags.Public | BindingFlags.Instance);
            method.Invoke(model, new object[] { tag });
        }
        catch (Exception ex)
        {
            Elmah.ErrorSignal.FromCurrentContext().Raise(ex);
        }
    }

    private void ReplaceTagsWithJobModel(string tag)
    {
        try
        {
            if (JobModel == null)
                return;

            switch (tag)
            {
                case "[JobId]":
                    ReplaceText(tag, this.JobId.ToString()); 
                    break;
                case "[JobStatus]":
                    ReplaceText(tag, JobModel.JobStatus);
                    break;
                case "[JobManager]":
                    ReplaceText(tag, JobModel.JobManager);
                    break;
                default:
                    break;
            }
        }
        catch (Exception ex)
        {
            Elmah.ErrorSignal.FromCurrentContext().Raise(ex);
        }
    }

    private void ReplaceTagsWithAttachmentModel(string tag)
    {
        try
        {
            if (JobModel == null)
                return;

            var attachmentId = JobModel.JobAttachmentId;
            if (attachmentId == 0)
                return;

            var attachmentModel = new AttachmentModel() { AttachmentId = attachmentId, JobId = this.JobId };
            switch (tag)
            {
                case "[AttachmentName]":
                    ReplaceText(tag, attachmentModel.AttachmentName);
                    break;
                case "[AttachmentPath]":
                    ReplaceText(tag, attachmentModel.AttachmentPath.Trim());
                    break;
                default:
                    break;
            }
        }
        catch (Exception ex)
        {
            Elmah.ErrorSignal.FromCurrentContext().Raise(ex);
        }
    }

    private void ReplaceText(string newText, string oldText)
    {
        try
        {
            if (ReportDoc == null)
                return;

            ReportDoc.Replace(newText, oldText, false, true, true);
        }
        catch (Exception ex)
        {
            Elmah.ErrorSignal.FromCurrentContext().Raise(ex);
        }
    }

    private byte[] SaveReport()
    {
        try
        {
            using(var ms = new MemoryStream())
            {
                ReportDoc.Save(ms, FormatType.Docx);
                ReportDoc.Close();

                return ms.ToArray();
            }
        }
        catch (Exception ex)
        {
            Elmah.ErrorSignal.FromCurrentContext().Raise(ex);
            return null;
        }
    }
}
