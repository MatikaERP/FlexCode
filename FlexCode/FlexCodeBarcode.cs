using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;


namespace FlexCode
{
    public class FlexCodeBarcode
    {
        public FlexCodeBarcode()
        {
            //Constructor Por Defecto
        }

        public static FlexRespuesta GeneratePDF417(string filepath, string data, int ecl)
        {
            var retorno = new FlexRespuesta();

            try
            {
                BarcodePDF417 barcode = new BarcodePDF417();
                barcode.Options = BarcodePDF417.PDF417_USE_ASPECT_RATIO;
                barcode.ErrorLevel = ecl;
                barcode.SetText(data);

                System.Drawing.Bitmap Imapdf417 = new Bitmap(barcode.CreateDrawingImage(System.Drawing.Color.Black, System.Drawing.Color.White));
                //PATH EN DONDE VA A SER GRABADO
                string actualdir = AppDomain.CurrentDomain.BaseDirectory;
                if (!Path.IsPathRooted(@filepath))
                {
                    filepath = actualdir + filepath;
                }

                //FileStream file = new FileStream(@filepath, FileMode.CreateNew);


                Imapdf417.Save(filepath, System.Drawing.Imaging.ImageFormat.Png);

                // file.Flush();
                //    file.Close();

                retorno.esCorrecto = true;
                retorno.mensaje = "Acción completada correctamente";

                return retorno;
            }
            catch (Exception e)
            {
                retorno.esCorrecto = false;
                retorno.mensaje = "Error:" + e.Message;

                return retorno;

            }
        }

    }

    public class FlexRespuesta
    {
        public bool esCorrecto;
        public string mensaje;
    }

    public class FlexEmailClient
    {
        public FlexEmailClient()
        {
            //Constructor Por Defecto
        }

        public static void IMAPReceiveEmails(string IMAPAddress, int Port, string User, string Password, Boolean SSL, String SavingPath)
        {
            try
            {
                Limilabs.Client.IMAP.Imap imap = new Limilabs.Client.IMAP.Imap();
                if (SSL == true)
                {
                    imap.ConnectSSL(IMAPAddress, Port);   // or ConnectSSL for SSL
                }
                else
                {
                    imap.Connect(IMAPAddress, Port);   // or ConnectSSL for SSL
                }

                imap.UseBestLogin(User, Password);
                imap.SelectInbox();

                List<long> uids = imap.Search(Limilabs.Client.IMAP.Flag.Unseen);

                foreach (long uid in uids)
                {
                    Limilabs.Mail.IMail email = new Limilabs.Mail.MailBuilder()
                        .CreateFromEml(imap.GetMessageByUID(uid));

                    Console.WriteLine(email.Subject);

                    // save all attachments to disk
                    foreach (Limilabs.Mail.MIME.MimeData mime in email.Attachments)
                    {

                        string fullName = Path.Combine(SavingPath, mime.SafeFileName);

                        //FileStream file = new FileStream(fullName, FileMode.Create);

                        byte[] data = mime.Data;

                        File.WriteAllBytes(fullName, data);

                        //mime.Save(file);

                        //file.Close();

                    }

                    imap.MarkMessageSeenByUID(uid);
                }
                imap.Close();
            }
            catch (Exception e)
            {
                var output = e.Output();
            }


        }

        public static void POP3ReceiveEmails(string POP3Address, int Port, string User, string Password, String SavingPath)
        {
            // C#
            Limilabs.Client.POP3.Pop3 pop3 = new Limilabs.Client.POP3.Pop3();
            pop3.ConnectSSL(POP3Address, Port);   // or ConnectSSL for SSL

            pop3.UseBestLogin(User, Password);
            foreach (string uid in pop3.GetAll())
            {
                Limilabs.Mail.IMail email = new Limilabs.Mail.MailBuilder()
                    .CreateFromEml(pop3.GetMessageByUID(uid));

                Console.WriteLine(email.Subject);

                // save all attachments to disk
                foreach (Limilabs.Mail.MIME.MimeData mime in email.Attachments)
                {
                    mime.Save(mime.SafeFileName);
                }
            }
            pop3.Close();
        }





    }


    public static class Extensions
    {
        public static string Output(this Exception ex)
        {
            if (ex == null) return String.Empty;

            var res = new StringBuilder();
            res.AppendFormat("Exception of type '{0}': {1}", ex.GetType().Name, ex.Message);
            res.AppendLine();

            if (!String.IsNullOrEmpty(ex.StackTrace))
            {
                res.AppendLine(ex.StackTrace);
            }

            if (ex.InnerException != null)
            {
                res.AppendLine(ex.InnerException.Output());
            }

            return res.ToString();
        }
    }
}
