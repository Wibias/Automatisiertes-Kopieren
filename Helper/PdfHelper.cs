using iText.Forms;
using iText.Kernel.Pdf;
using System;
using System.IO;
using System.Threading.Tasks;
using static Automatisiertes_Kopieren.Helper.LoggingHelper;

namespace Automatisiertes_Kopieren.Helper
{
    public static class PdfHelper
    {
        public enum PdfType
        {
            Protokollbogen,
            AllgemeinEntwicklungsbericht,
            ProtokollElterngespraech,
            VorschuleEntwicklungsbericht,
            KrippeUebergangsbericht
        }

        public static async Task FillPdfAsync(string pdfPath, string kidName, double monthsValue, string group,
            PdfType pdfType, string parsedBirthDate, string? genderValue)
        {
            string tempPath = pdfPath + ".temp";

            try
            {
                using (var pdfDoc = new PdfDocument(new PdfReader(pdfPath), new PdfWriter(tempPath)))
                {
                    var form = PdfAcroForm.GetAcroForm(pdfDoc, true) ??
                               throw new Exception("Das PDF enthält keine Formularfelder.");

                    FillPdfForm(form, pdfType, kidName, monthsValue, group, parsedBirthDate, genderValue);

                    pdfDoc.Close();
                }

                await ReplaceOriginalFileAsync(pdfPath, tempPath).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogException(ex, $"Fehler in FillPdfAsync aufgetreten. {ex.Message}");
                await CleanupTempFileAsync(tempPath).ConfigureAwait(false);

            }
        }

        private static void FillPdfForm(PdfAcroForm form, PdfType pdfType, string kidName, double monthsValue, string group,
            string parsedBirthDate, string? genderValue)
        {
            switch (pdfType)
            {
                case PdfType.Protokollbogen:
                    form.GetField("Name_des_Kindes").SetValue(kidName);
                    form.GetField("Alter_des_Kindes_in_Monaten").SetValue(monthsValue.ToString("0.00"));
                    form.GetField("Gruppe").SetValue(group);
                    form.GetField("Heutiges_Datum").SetValue(DateTime.Now.ToString("dd.MM.yyyy"));
                    form.GetField("Geburtsdatum").SetValue(parsedBirthDate);

                    form.GetField("männlich").SetValue(genderValue == "Männlich" ? "On" : "Off");
                    form.GetField("weiblich").SetValue(genderValue == "Weiblich" ? "On" : "Off");
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

                case PdfType.VorschuleEntwicklungsbericht:
                    form.GetField("Name des Kindes").SetValue(kidName);
                    form.GetField("Datum").SetValue(DateTime.Now.ToString("dd.MM.yyyy"));
                    form.GetField("Gruppe").SetValue(group);
                    break;

                case PdfType.KrippeUebergangsbericht:
                    form.GetField("Name des Kindes").SetValue(kidName);
                    form.GetField("Datum").SetValue(DateTime.Now.ToString("dd.MM.yyyy"));
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(pdfType), pdfType, null);
            }
        }

        private static async Task ReplaceOriginalFileAsync(string originalPath, string tempPath)
        {
            try
            {
                if (File.Exists(originalPath))
                {
                    using (var sourceStream = new FileStream(originalPath, FileMode.Truncate))
                    using (var destinationStream = new FileStream(tempPath, FileMode.Open, FileAccess.Read))
                    {
                        await destinationStream.CopyToAsync(sourceStream).ConfigureAwait(false);
                    }
                }

                await Task.Run(() => File.Delete(tempPath)).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                LogException(ex, $"Fehler beim Ersetzen der temporären Datei. {ex.Message}");
                await CleanupTempFileAsync(tempPath).ConfigureAwait(false);
            }
        }

        private static async Task CleanupTempFileAsync(string tempPath)
        {
            try
            {
                if (File.Exists(tempPath))
                {
                    using (var fileStream = new FileStream(tempPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.Asynchronous | FileOptions.DeleteOnClose))
                    {
                        await fileStream.FlushAsync().ConfigureAwait(false);
                    }
                }
            }
            catch (Exception ex)
            {
                LogException(ex, $"Fehler beim Löschen der temporären Datei. {ex.Message}");
            }
        }
     }
}