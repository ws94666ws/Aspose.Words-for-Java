package DocsExamples.File_formats_and_conversions.Save_options;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

@Test
public class WorkingWithOoxmlSaveOptions extends DocsExamplesBase {
    @Test
    public void encryptDocxWithPassword() throws Exception {
        //ExStart:EncryptDocxWithPassword
        Document doc = new Document(getMyDir() + "Document.docx");

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.setPassword("password");

        doc.save(getArtifactsDir() + "WorkingWithOoxmlSaveOptions.EncryptDocxWithPassword.docx", saveOptions);
        //ExEnd:EncryptDocxWithPassword
    }

    @Test
    public void ooxmlComplianceIso29500_2008_Strict() throws Exception {
        //ExStart:OoxmlComplianceIso29500_2008_Strict
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2016);

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.setCompliance(OoxmlCompliance.ISO_29500_2008_STRICT);

        doc.save(getArtifactsDir() + "WorkingWithOoxmlSaveOptions.OoxmlComplianceIso29500_2008_Strict.docx", saveOptions);
        //ExEnd:OoxmlComplianceIso29500_2008_Strict
    }

    @Test
    public void updateLastSavedTime() throws Exception {
        //ExStart:UpdateLastSavedTime
        //GistId:a6f7799aa265589fb56915bb1e401b05
        Document doc = new Document(getMyDir() + "Document.docx");

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.setUpdateLastSavedTimeProperty(true);

        doc.save(getArtifactsDir() + "WorkingWithOoxmlSaveOptions.UpdateLastSavedTime.docx", saveOptions);
        //ExEnd:UpdateLastSavedTime
    }

    @Test
    public void keepLegacyControlChars() throws Exception {
        //ExStart:KeepLegacyControlChars
        Document doc = new Document(getMyDir() + "Legacy control character.doc");

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.FLAT_OPC);
        saveOptions.setKeepLegacyControlChars(true);

        doc.save(getArtifactsDir() + "WorkingWithOoxmlSaveOptions.KeepLegacyControlChars.docx", saveOptions);
        //ExEnd:KeepLegacyControlChars
    }

    @Test
    public void setCompressionLevel() throws Exception {
        //ExStart:SetCompressionLevel
        Document doc = new Document(getMyDir() + "Document.docx");

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.setCompressionLevel(CompressionLevel.SUPER_FAST);

        doc.save(getArtifactsDir() + "WorkingWithOoxmlSaveOptions.SetCompressionLevel.docx", saveOptions);
        //ExEnd:SetCompressionLevel
    }
}
