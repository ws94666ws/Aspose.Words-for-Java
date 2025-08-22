package DocsExamples.File_formats_and_conversions.Save_options;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

import java.util.ArrayList;
import java.util.Date;

@Test
public class WorkingWithPdfSaveOptions extends DocsExamplesBase {
    @Test
    public void displayDocTitleInWindowTitleBar() throws Exception {
        //ExStart:DisplayDocTitleInWindowTitleBar
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setDisplayDocTitle(true);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.DisplayDocTitleInWindowTitlebar.pdf", saveOptions);
        //ExEnd:DisplayDocTitleInWindowTitleBar
    }

    @Test
    //ExStart:PdfRenderWarnings
    //GistId:c33834c88b84242b9b28c1cfc22eb762
    public void pdfRenderWarnings() throws Exception {
        Document doc = new Document(getMyDir() + "WMF with image.docx");

        MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();
        metafileRenderingOptions.setEmulateRasterOperations(false);
        metafileRenderingOptions.setRenderingMode(MetafileRenderingMode.VECTOR_WITH_FALLBACK);

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setMetafileRenderingOptions(metafileRenderingOptions);

        // If Aspose.Words cannot correctly render some of the metafile records
        // to vector graphics then Aspose.Words renders this metafile to a bitmap.
        HandleDocumentWarnings callback = new HandleDocumentWarnings();
        doc.setWarningCallback(callback);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.PdfRenderWarnings.pdf", saveOptions);

        // While the file saves successfully, rendering warnings that occurred during saving are collected here.
        for (WarningInfo warningInfo : callback.mWarnings) {
            System.out.println(warningInfo.getDescription());
        }
    }

    public static class HandleDocumentWarnings implements IWarningCallback {
        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during
        /// document load and/or document save.
        /// </summary>
        public void warning(WarningInfo info) {
            // For now type of warnings about unsupported metafile records changed
            // from DataLoss/UnexpectedContent to MinorFormattingLoss.
            if (info.getWarningType() == WarningType.MINOR_FORMATTING_LOSS) {
                System.out.println("Unsupported operation: " + info.getDescription());
                mWarnings.warning(info);
            }
        }

        public WarningInfoCollection mWarnings = new WarningInfoCollection();
    }
    //ExEnd:PdfRenderWarnings

    @Test
    public void digitallySignedPdfUsingCertificateHolder() throws Exception {
        //ExStart:DigitallySignedPdfUsingCertificateHolder
        //GistId:39ea49b7754e472caf41179f8b5970a0
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Test Signed PDF.");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setDigitalSignatureDetails(new PdfDigitalSignatureDetails(
                CertificateHolder.create(getMyDir() + "morzal.pfx", "aw"), "reason", "location",
                new Date()));

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.DigitallySignedPdfUsingCertificateHolder.pdf", saveOptions);
        //ExEnd:DigitallySignedPdfUsingCertificateHolder
    }

    @Test
    public void embeddedAllFonts() throws Exception {
        //ExStart:EmbeddedAllFonts
        //GistId:a5d65fc091d4330c8b66a17170524341
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // The output PDF will be embedded with all fonts found in the document.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setEmbedFullFonts(true);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.EmbeddedAllFonts.pdf", saveOptions);
        //ExEnd:EmbeddedAllFonts
    }

    @Test
    public void embeddedSubsetFonts() throws Exception {
        //ExStart:EmbeddedSubsetFonts
        //GistId:a5d65fc091d4330c8b66a17170524341
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // The output PDF will contain subsets of the fonts in the document.
        // Only the glyphs used in the document are included in the PDF fonts.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setEmbedFullFonts(false);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.EmbeddedSubsetFonts.pdf", saveOptions);
        //ExEnd:EmbeddedSubsetFonts
    }

    @Test
    public void disableEmbedWindowsFonts() throws Exception {
        //ExStart:DisableEmbedWindowsFonts
        //GistId:a5d65fc091d4330c8b66a17170524341
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // The output PDF will be saved without embedding standard windows fonts.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setFontEmbeddingMode(PdfFontEmbeddingMode.EMBED_NONE);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.DisableEmbedWindowsFonts.pdf", saveOptions);
        //ExEnd:DisableEmbedWindowsFonts
    }

    @Test
    public void skipEmbeddedArialAndTimesRomanFonts() throws Exception {
        //ExStart:SkipEmbeddedArialAndTimesRomanFonts
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setFontEmbeddingMode(PdfFontEmbeddingMode.EMBED_ALL);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.SkipEmbeddedArialAndTimesRomanFonts.pdf", saveOptions);
        //ExEnd:SkipEmbeddedArialAndTimesRomanFonts
    }

    @Test
    public void avoidEmbeddingCoreFonts() throws Exception {
        //ExStart:AvoidEmbeddingCoreFonts
        //GistId:a5d65fc091d4330c8b66a17170524341
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setUseCoreFonts(true);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.AvoidEmbeddingCoreFonts.pdf", saveOptions);
        //ExEnd:AvoidEmbeddingCoreFonts
    }

    @Test
    public void escapeUri() throws Exception {
        //ExStart:EscapeUri
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertHyperlink("Testlink", "https://www.google.com/search?q= aspose", false);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.EscapeUri.pdf");
        //ExEnd:EscapeUri
    }

    @Test
    public void exportHeaderFooterBookmarks() throws Exception {
        //ExStart:ExportHeaderFooterBookmarks
        //GistId:a5d65fc091d4330c8b66a17170524341
        Document doc = new Document(getMyDir() + "Bookmarks in headers and footers.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.getOutlineOptions().setDefaultBookmarksOutlineLevel(1);
        saveOptions.setHeaderFooterBookmarksExportMode(HeaderFooterBookmarksExportMode.FIRST);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.ExportHeaderFooterBookmarks.pdf", saveOptions);
        //ExEnd:ExportHeaderFooterBookmarks
    }

    @Test
    public void emulateRenderingToSizeOnPage() throws Exception {
        //ExStart:EmulateRenderingToSizeOnPage
        Document doc = new Document(getMyDir() + "WMF with text.docx");

        MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();
        metafileRenderingOptions.setEmulateRenderingToSizeOnPage(false);

        // If Aspose.Words cannot correctly render some of the metafile records to vector graphics
        // then Aspose.Words renders this metafile to a bitmap.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setMetafileRenderingOptions(metafileRenderingOptions);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.EmulateRenderingToSizeOnPage.pdf", saveOptions);
        //ExEnd:EmulateRenderingToSizeOnPage
    }

    @Test
    public void additionalTextPositioning() throws Exception {
        //ExStart:AdditionalTextPositioning
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setAdditionalTextPositioning(true);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
        //ExEnd:AdditionalTextPositioning
    }

    @Test
    public void conversionToPdf17() throws Exception {
        //ExStart:ConversionToPdf17
        //GistId:b237846932dfcde42358bd0c887661a5
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setCompliance(PdfCompliance.PDF_17);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.ConversionToPdf17.pdf", saveOptions);
        //ExEnd:ConversionToPdf17
    }

    @Test
    public void downsamplingImages() throws Exception {
        //ExStart:DownsamplingImages
        //GistId:a5d65fc091d4330c8b66a17170524341
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // We can set a minimum threshold for downsampling.
        // This value will prevent the second image in the input document from being downsampled.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.getDownsampleOptions().setResolution(36);
        saveOptions.getDownsampleOptions().setResolutionThreshold(128);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.DownsamplingImages.pdf", saveOptions);
        //ExEnd:DownsamplingImages
    }

    @Test
    public void outlineOptions() throws Exception {
        //ExStart:OutlineOptions
        //GistId:a5d65fc091d4330c8b66a17170524341
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.getOutlineOptions().setHeadingsOutlineLevels(3);
        saveOptions.getOutlineOptions().setExpandedOutlineLevels(1);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.OutlineOptions.pdf", saveOptions);
        //ExEnd:OutlineOptions
    }

    @Test
    public void customPropertiesExport() throws Exception {
        //ExStart:CustomPropertiesExport
        //GistId:a5d65fc091d4330c8b66a17170524341
        Document doc = new Document();
        doc.getCustomDocumentProperties().add("Company", "Aspose");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setCustomPropertiesExport(PdfCustomPropertiesExport.STANDARD);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.CustomPropertiesExport.pdf", saveOptions);
        //ExEnd:CustomPropertiesExport
    }

    @Test
    public void exportDocumentStructure() throws Exception {
        //ExStart:ExportDocumentStructure
        //GistId:a5d65fc091d4330c8b66a17170524341
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // The file size will be increased and the structure will be visible in the "Content" navigation pane
        // of Adobe Acrobat Pro, while editing the .pdf.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setExportDocumentStructure(true);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.ExportDocumentStructure.pdf", saveOptions);
        //ExEnd:ExportDocumentStructure
    }

    @Test
    public void imageCompression() throws Exception {
        //ExStart:ImageCompression
        //GistId:a5d65fc091d4330c8b66a17170524341
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setImageCompression(PdfImageCompression.JPEG);
        saveOptions.setPreserveFormFields(true);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.ImageCompression.pdf", saveOptions);

        PdfSaveOptions saveOptionsA2U = new PdfSaveOptions();
        saveOptionsA2U.setCompliance(PdfCompliance.PDF_A_2_U);
        saveOptionsA2U.setImageCompression(PdfImageCompression.JPEG);
        saveOptionsA2U.setJpegQuality(100); // Use JPEG compression at 50% quality to reduce file size.

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.ImageCompression_A2u.pdf", saveOptionsA2U);
        //ExEnd:ImageCompression
    }

    @Test
    public void updateLastPrinted() throws Exception {
        //ExStart:UpdateLastPrinted
        //GistId:a6f7799aa265589fb56915bb1e401b05
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setUpdateLastPrintedProperty(true);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.UpdateLastPrinted.pdf", saveOptions);
        //ExEnd:UpdateLastPrinted
    }

    @Test
    public void dml3DEffectsRendering() throws Exception {
        //ExStart:Dml3DEffectsRendering
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setDml3DEffectsRenderingMode(Dml3DEffectsRenderingMode.ADVANCED);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.Dml3DEffectsRendering.pdf", saveOptions);
        //ExEnd:Dml3DEffectsRendering
    }

    @Test
    public void interpolateImages() throws Exception {
        //ExStart:SetImageInterpolation
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setInterpolateImages(true);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.InterpolateImages.pdf", saveOptions);
        //ExEnd:SetImageInterpolation
    }

    @Test
    public void optimizeOutput() throws Exception {
        //ExStart:OptimizeOutput
        //GistId:b237846932dfcde42358bd0c887661a5
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setOptimizeOutput(true);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.OptimizeOutput.pdf", saveOptions);
        //ExEnd:OptimizeOutput
    }

    @Test
    public void updateScreenTip() throws Exception {
        //ExStart:UpdateScreenTip
        //GistId:d92cef7ddb3b69b3f59a83b0c749326d
        Document doc = new Document(getMyDir() + "Table of contents.docx");

        // Get all hyperlink fields that are TOC links (SubAddress starts with "#_Toc").
        FieldCollection fields = doc.getRange().getFields();
        ArrayList<FieldHyperlink> tocHyperLinks = new ArrayList<>();

        for (Field field : fields) {
            if (field.getType() == FieldType.FIELD_HYPERLINK) {
                FieldHyperlink hyperlink = (FieldHyperlink) field;
                if (hyperlink.getSubAddress() != null && hyperlink.getSubAddress().startsWith("#_Toc")) {
                    tocHyperLinks.add(hyperlink);
                }
            }
        }

        // Update ScreenTip for each TOC hyperlink
        for (FieldHyperlink link : tocHyperLinks) {
            link.setScreenTip(link.getDisplayResult());
        }

        // Configure PDF save options
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setCompliance(PdfCompliance.PDF_UA_1);
        saveOptions.setDisplayDocTitle(true);
        saveOptions.setExportDocumentStructure(true);

        // Configure outline options
        saveOptions.getOutlineOptions().setHeadingsOutlineLevels(3);
        saveOptions.getOutlineOptions().setCreateMissingOutlineLevels(true);

        doc.save(getArtifactsDir() + "WorkingWithPdfSaveOptions.UpdateScreenTip.pdf", saveOptions);
        //ExEnd:UpdateScreenTip
    }
}
