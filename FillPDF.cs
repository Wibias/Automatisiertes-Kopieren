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
        public enum PdfType
        {
            Protokollbogen,
            AllgemeinEntwicklungsbericht,
            ProtokollElterngespraech,
            VorschulEntwicklungsbericht
        }

        public void FillPdf(string pdfPath, string kidName, double monthsValue, string group, PdfType pdfType, string parsedBirthDate, string? genderValue)
        {
            try
            {
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(pdfPath), new PdfWriter(pdfPath + ".temp"));

                PdfAcroForm form = PdfAcroForm.GetAcroForm(pdfDoc, true);

                if (form == null)
                {
                    throw new Exception("The PDF does not contain any form fields.");
                }

                switch (pdfType)
                {
                    case PdfType.Protokollbogen:
                        form.GetField("Name_des_Kindes").SetValue(kidName);
                        form.GetField("Alter_des_Kindes_in_Monaten").SetValue(monthsValue.ToString("0.00"));
                        form.GetField("Gruppe").SetValue(group);
                        form.GetField("Heutiges_Datum").SetValue(DateTime.Now.ToString("dd.MM.yyyy"));
                        form.GetField("Geburtsdatum").SetValue(parsedBirthDate);

                        if (genderValue == "Männlich")
                        {
                            form.GetField("männlich").SetValue("On");
                            form.GetField("weiblich").SetValue("Off");
                        }
                        else if (genderValue == "Weiblich")
                        {
                            form.GetField("weiblich").SetValue("On");
                            form.GetField("männlich").SetValue("Off");
                        }
                        break;

                    case PdfType.AllgemeinEntwicklungsbericht:
                        form.GetField("Name").SetValue(kidName);
                        form.GetField("Alter in Monaten").SetValue(monthsValue.ToString("0.00"));
                        form.GetField("Gruppe").SetValue(group);
                        form.GetField("Datum").SetValue(DateTime.Now.ToString("dd.MM.yyyy"));
                        break;

                    case PdfType.ProtokollElterngespraech:
                        form.GetField("Name des Kindes").SetValue(kidName);
                        form.GetField("Geburtsdatum").SetValue(parsedBirthDate);
                        break;
                    case PdfType.VorschulEntwicklungsbericht:
                        form.GetField("Name des Kindes").SetValue(kidName);
                        form.GetField("Datum").SetValue(DateTime.Now.ToString("dd.MM.yyyy"));
                        form.GetField("Gruppe").SetValue(group);
                        break;
                }

                pdfDoc.Close();
            }
            catch (Exception ex)
            {
                _loggingService.LogMessage($"Error encountered in FillPdf. Message: {ex.Message}. StackTrace: {ex.StackTrace}", LogLevel.Error);
            }
            try
            {
                File.Delete(pdfPath);
                File.Move(pdfPath + ".temp", pdfPath);
            }
            catch (Exception ex)
            {
                _loggingService.LogMessage($"Error encountered while handling file operations. Message: {ex.Message}. StackTrace: {ex.StackTrace}", LogLevel.Error);
            }
        }
    }
}