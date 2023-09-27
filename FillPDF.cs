using iText.Forms;
using iText.Kernel.Pdf;
using System;
using System.IO;
using static Automatisiertes_Kopieren.LoggingService;

namespace Automatisiertes_Kopieren
{
    public class FillPDF
    {
        private readonly static LoggingService _loggingService = new LoggingService();

        public void FillProtokollbogen(string renamedProtokollbogenPath, string kidName, double monthsValue, string group)
        {
            try
            {
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(renamedProtokollbogenPath), new PdfWriter(renamedProtokollbogenPath + ".temp"));

                PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);

                if (form == null)
                {
                    throw new Exception("The PDF does not contain any form fields.");
                }

                form.GetField("Name_des_Kindes").SetValue(kidName);
                form.GetField("Alter_des_Kindes_in_Monaten").SetValue(monthsValue.ToString("0.00"));
                form.GetField("Gruppe").SetValue(group);

                string currentDate = DateTime.Now.ToString("dd.MM.yyyy");
                form.GetField("Heutiges_Datum").SetValue(currentDate);

                pdfDoc.Close();
            }
            catch (Exception ex)
            {
                _loggingService.LogMessage($"Error encountered in FillProtokollbogen. Message: {ex.Message}. StackTrace: {ex.StackTrace}", LogLevel.Error);
            }

            try
            {
                File.Delete(renamedProtokollbogenPath);
                File.Move(renamedProtokollbogenPath + ".temp", renamedProtokollbogenPath);
            }
            catch (Exception ex)
            {
                _loggingService.LogMessage($"Error encountered while handling file operations. Message: {ex.Message}. StackTrace: {ex.StackTrace}", LogLevel.Error);
            }
        }

    }
}