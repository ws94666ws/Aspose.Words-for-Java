package DocsExamples.File_formats_and_conversions.Save_options;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Document;
import com.aspose.words.PclSaveOptions;
import com.aspose.words.SaveFormat;
import org.testng.annotations.Test;

@Test
public class WorkingWithPclSaveOptions extends DocsExamplesBase {
    @Test
    public void rasterizeTransformedElements() throws Exception {
        //ExStart:RasterizeTransformedElements
        //GistId:9d2a393f6dff9d785e7747a48e590d9d
        Document doc = new Document(getMyDir() + "Rendering.docx");

        PclSaveOptions saveOptions = new PclSaveOptions();
        saveOptions.setSaveFormat(SaveFormat.PCL);
        saveOptions.setRasterizeTransformedElements(false);

        doc.save(getArtifactsDir() + "WorkingWithPclSaveOptions.RasterizeTransformedElements.pcl", saveOptions);
        //ExEnd:RasterizeTransformedElements
    }
}
