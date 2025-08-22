package DocsExamples.Programming_with_documents.Contents_management;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Document;
import org.testng.annotations.Test;

@Test
public class WorkingWithRanges extends DocsExamplesBase {
    @Test
    public void rangesDeleteText() throws Exception {
        //ExStart:RangesDeleteText
        //GistId:f3b385f86a13b93093c214b1e7a981bf
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.getSections().get(0).getRange().delete();
        //ExEnd:RangesDeleteText
    }

    @Test
    public void rangesGetText() throws Exception {
        //ExStart:RangesGetText
        //GistId:f3b385f86a13b93093c214b1e7a981bf
        Document doc = new Document(getMyDir() + "Document.docx");
        String text = doc.getRange().getText();
        //ExEnd:RangesGetText
    }
}
