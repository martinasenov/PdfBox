package Bookmark.Nomad;

import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotationWidget;
import org.apache.pdfbox.pdmodel.interactive.form.PDAcroForm;
import org.apache.pdfbox.pdmodel.interactive.form.PDTextField;

import java.io.File;
import java.io.IOException;

public class AddTextFieldToPDF {

    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\martin.asenov\\Downloads\\WORKORDERTASKCARD_39280475_FLATTENED.pdf";
        String outputFilePath = inputFilePath+"_OUTPUT.pdf";

        try {
            PDDocument document = PDDocument.load(new File(inputFilePath));
            PDAcroForm acroForm = document.getDocumentCatalog().getAcroForm();

            if (acroForm == null) {
                acroForm = new PDAcroForm(document);
                document.getDocumentCatalog().setAcroForm(acroForm);
            }

            // Set the default resources for the AcroForm
            PDResources dr = new PDResources();
            acroForm.setDefaultResources(dr);

            // Adding the default font to the resources
            dr.put(COSName.getPDFName("Helv-Bold"), PDType1Font.HELVETICA_BOLD);

            PDPageTree pages = document.getDocumentCatalog().getPages();

            int pageIndex = 0;
            for (PDPage page : pages) {
                // First Text Field: Maint.Notes
                PDTextField textField1 = new PDTextField(acroForm);
                textField1.setPartialName("Maintenance Notes" + pageIndex);
                PDAnnotationWidget widget1 = textField1.getWidgets().get(0);
                PDRectangle rect1 = new PDRectangle(51, 195, 478, 221);
                widget1.setRectangle(rect1);
                textField1.setDefaultAppearance("/Helv-Bold 0 Tf 0 g");
                textField1.setMultiline(true);
                acroForm.getFields().add(textField1);
                page.getAnnotations().add(widget1);

                // Second Text Field: TECHNICIAN SIGNATURE
                PDTextField textField2 = new PDTextField(acroForm);
                textField2.setPartialName("TECHNICIAN SIGNATURE" + pageIndex);
                PDAnnotationWidget widget2 = textField2.getWidgets().get(0);
                PDRectangle rect2 = new PDRectangle(51, 120, 156, 18);
                widget2.setRectangle(rect2);
                textField2.setDefaultAppearance("/Helv-Bold 0 Tf 0 g");
                acroForm.getFields().add(textField2);
                page.getAnnotations().add(widget2);

                // Third Text Field: INSPECTED BY
                PDTextField textField3 = new PDTextField(acroForm);
                textField3.setPartialName("INSPECTED BY_" + pageIndex);
                PDAnnotationWidget widget3 = textField3.getWidgets().get(0);
                PDRectangle rect3 = new PDRectangle(51, 85, 156, 18);
                widget3.setRectangle(rect3);
                textField3.setDefaultAppearance("/Helv-Bold 0 Tf 0 g");
                acroForm.getFields().add(textField3);
                page.getAnnotations().add(widget3);


                // Third Text Field: CERTIFICATE
                PDTextField textField4 = new PDTextField(acroForm);
                textField4.setPartialName("CERTIFICATE1_" + pageIndex);
                PDAnnotationWidget widget4 = textField4.getWidgets().get(0);
                PDRectangle rect4 = new PDRectangle(210, 132, 156, 11.5F);
                widget4.setRectangle(rect4);
                textField4.setDefaultAppearance("/Helv-Bold 0 Tf 0 g");
                textField4.setQ(1);
                acroForm.getFields().add(textField4);
                page.getAnnotations().add(widget4);

                // Third Text Field: STAMP1
                PDTextField textField5 = new PDTextField(acroForm);
                textField5.setPartialName("STAMP1_" + pageIndex);
                PDAnnotationWidget widget5 = textField5.getWidgets().get(0);
                PDRectangle rect5 = new PDRectangle(210, 120, 156, 11.5f);
                widget5.setRectangle(rect5);
                textField5.setDefaultAppearance("/Helv-Bold 0 Tf 0 g");
                textField5.setQ(1);
                acroForm.getFields().add(textField5);
                page.getAnnotations().add(widget5);



                // Third Text Field: DATE1
                PDTextField textField6 = new PDTextField(acroForm);
                textField6.setPartialName("DATE1_" + pageIndex);
                PDAnnotationWidget widget6 = textField6.getWidgets().get(0);
                PDRectangle rect6 = new PDRectangle(370, 120, 156, 22);
                widget6.setRectangle(rect6);
                textField6.setDefaultAppearance("/Helv-Bold 0 Tf 0 g");
                acroForm.getFields().add(textField6);
                page.getAnnotations().add(widget6);


                // Third Text Field: DATE2
                PDTextField textField7 = new PDTextField(acroForm);
                textField7.setPartialName("DATE2_" + pageIndex);
                PDAnnotationWidget widget7 = textField7.getWidgets().get(0);
                PDRectangle rect7 = new PDRectangle(370, 84, 156, 22);
                widget7.setRectangle(rect7);
                textField7.setDefaultAppearance("/Helv-Bold 0 Tf 0 g");
                acroForm.getFields().add(textField7);
                page.getAnnotations().add(widget7);

                // Third Text Field: CERTIFICATE2
                PDTextField textField8 = new PDTextField(acroForm);
                textField8.setPartialName("CERTIFICATE2_" + pageIndex);
                PDAnnotationWidget widget8 = textField8.getWidgets().get(0);
                PDRectangle rect8 = new PDRectangle(210, 97, 156, 11.5F);
                widget8.setRectangle(rect8);
                textField8.setDefaultAppearance("/Helv-Bold 0 Tf 0 g");
                textField8.setQ(1);
                acroForm.getFields().add(textField8);
                page.getAnnotations().add(widget8);

                // Third Text Field: STAMP3
                PDTextField textField9 = new PDTextField(acroForm);
                textField9.setPartialName("STAMP3_" + pageIndex);
                PDAnnotationWidget widget9 = textField9.getWidgets().get(0);
                PDRectangle rect9 = new PDRectangle(210, 85, 156, 11.5F);
                widget9.setRectangle(rect9);
                textField9.setDefaultAppearance("/Helv-Bold 0 Tf 0 g");
                textField9.setQ(1);
                acroForm.getFields().add(textField9);
                page.getAnnotations().add(widget9);




                pageIndex++;
            }

            document.save(new File(outputFilePath));
            document.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}




