package DocsExamples.File_formats_and_conversions.Save_options;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

import java.io.ByteArrayOutputStream;

@Test
public class WorkingWithMarkdownSaveOptions extends DocsExamplesBase {
    @Test
    public void markdownTableContentAlignment() throws Exception {
        //ExStart:MarkdownTableContentAlignment
        //GistId:50b2b6a8785c07713e7c09d772e9a396
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertCell();
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
        builder.write("Cell1");
        builder.insertCell();
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.write("Cell2");

        // Makes all paragraphs inside the table to be aligned.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        saveOptions.setTableContentAlignment(TableContentAlignment.LEFT);

        doc.save(getArtifactsDir() + "WorkingWithMarkdownSaveOptions.LeftTableContentAlignment.md", saveOptions);

        saveOptions.setTableContentAlignment(TableContentAlignment.RIGHT);
        doc.save(getArtifactsDir() + "WorkingWithMarkdownSaveOptions.RightTableContentAlignment.md", saveOptions);

        saveOptions.setTableContentAlignment(TableContentAlignment.CENTER);
        doc.save(getArtifactsDir() + "WorkingWithMarkdownSaveOptions.CenterTableContentAlignment.md", saveOptions);

        // The alignment in this case will be taken from the first paragraph in corresponding table column.
        saveOptions.setTableContentAlignment(TableContentAlignment.AUTO);
        doc.save(getArtifactsDir() + "WorkingWithMarkdownSaveOptions.AutoTableContentAlignment.md", saveOptions);
        //ExEnd:MarkdownTableContentAlignment
    }

    @Test
    public void imagesFolder() throws Exception {
        //ExStart:ImagesFolder
        //GistId:642767bbe8d8bec8eab080120b707990
        Document doc = new Document(getMyDir() + "Image bullet points.docx");

        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        saveOptions.setImagesFolder(getArtifactsDir() + "Images");

        try (ByteArrayOutputStream stream = new ByteArrayOutputStream()) {
            doc.save(stream, saveOptions);
        }
        //ExEnd:ImagesFolder
    }
}

