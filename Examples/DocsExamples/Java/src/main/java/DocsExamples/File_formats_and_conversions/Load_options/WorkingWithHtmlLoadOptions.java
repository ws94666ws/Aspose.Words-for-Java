package DocsExamples.File_formats_and_conversions.Load_options;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Document;
import com.aspose.words.HtmlControlType;
import com.aspose.words.HtmlLoadOptions;
import com.aspose.words.SaveFormat;
import org.testng.annotations.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;

@Test
public class WorkingWithHtmlLoadOptions extends DocsExamplesBase {
    @Test
    public void preferredControlType() throws Exception {
        //ExStart:LoadHtmlElementsWithPreferredControlType
        final String html = "<html>" +
                "<select name='ComboBox' size='1'>" +
                "<option value='val1'>item1</option>" +
                "<option value='val2'></option>" +
                "</select>" +
                "</html>";

        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        loadOptions.setPreferredControlType(HtmlControlType.STRUCTURED_DOCUMENT_TAG);

        Document doc = new Document(new ByteArrayInputStream(html.getBytes(StandardCharsets.UTF_8)), loadOptions);

        doc.save(getArtifactsDir() + "WorkingWithHtmlLoadOptions.PreferredControlType.docx", SaveFormat.DOCX);
        //ExEnd:LoadHtmlElementsWithPreferredControlType
    }
}
