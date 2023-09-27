using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.AcroForms;
using PdfSharpCore.Pdf.IO;

public class FillPDF
{
    public void FillProtokollbogen(string protokollbogenFilePath, string kidName, int monthsValue, string group)
    {
        // Open the Protokollbogen PDF file
        using (PdfDocument document = PdfReader.Open(protokollbogenFilePath, PdfDocumentOpenMode.Modify))
        {
            PdfPage page = document.Pages[0];

            XGraphics gfx = XGraphics.FromPdfPage(page);

            PdfAcroForm form = document.AcroForm;

            PdfTextField? nameTextField = form.Fields["Name_des_Kindes"] as PdfTextField;
            if (nameTextField != null)
            {
                // Create a PdfString to set the field value
                nameTextField.Value = new PdfString(kidName);
            }

            PdfTextField? monthsTextField = form.Fields["Alter_des_Kindes_in_Monaten"] as PdfTextField;
            if (monthsTextField != null)
            {
                monthsTextField.Value = new PdfString(monthsValue.ToString());
            }

            PdfTextField? groupTextField = form.Fields["Gruppe"] as PdfTextField;
            if (groupTextField != null)
            {
                groupTextField.Value = new PdfString(group);
            }

            document.Save(protokollbogenFilePath);
        }
    }
}