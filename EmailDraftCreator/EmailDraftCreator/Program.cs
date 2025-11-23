using System;
using System.Web;
using System.Windows.Forms;                                      // עבור OpenFileDialog
using Outlook = Microsoft.Office.Interop.Outlook;                // Alias כדי להימנע מהתנגשויות

namespace EmailDraftCreator
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                    return; // לא הגיע URL → לא עושים כלום

                string url = args[0];

                // ניתוח ה-URL
                Uri uri = new Uri(url);
                var query = HttpUtility.ParseQueryString(uri.Query);

                string subject = query["subject"] ?? "";
                string recipients = query["recipients"] ?? "";
                string body = query["body"] ?? "";
                string file = query["file"] ?? ""; // תמיד ריק מהדפדפן

                // אם לא הועבר קובץ → המשתמש בוחר קובץ
                if (string.IsNullOrEmpty(file))
                {
                    OpenFileDialog dlg = new OpenFileDialog();
                    dlg.Title = "בחר קובץ קורות חיים לצירוף";
                    dlg.Filter = "All Files (*.*)|*.*";

                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        file = dlg.FileName;
                    }
                }

                // יצירת טיוטות – אחת לכל נמען
                Outlook.Application outlook = new Outlook.Application();

                foreach (string recipient in recipients.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    Outlook.MailItem mail = (Outlook.MailItem)outlook.CreateItem(Outlook.OlItemType.olMailItem);

                    mail.Subject = subject;
                    mail.Body = body;
                    mail.To = recipient.Trim();

                    if (!string.IsNullOrWhiteSpace(file))
                        mail.Attachments.Add(file);

                    mail.Display(); // פותח טיוטה
                }
            }
            catch (UriFormatException ex)
            {
                System.IO.File.WriteAllText("error_log.txt", "URL Format Error:\n" + ex.ToString());
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                System.IO.File.WriteAllText("error_log.txt", "Outlook COM Error:\n" + ex.ToString());
            }
            catch (System.Exception ex)
            {
                System.IO.File.WriteAllText("error_log.txt", "General Error:\n" + ex.ToString());
            }
        }
    }
}