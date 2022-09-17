namespace EmailReseiver.Models
{
    public class MailItem
    {
        public string Date { get;  set; }
        public string From { get;  set; }
        public object Subj { get;  set; }
        public bool HasAttachments { get; set; }
    }
}
