package DocsExamples.File_formats_and_conversions.Save_options;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

import java.text.MessageFormat;

@Test
public class WorkingWithImageSaveOptions extends DocsExamplesBase {
    @Test
    public void exposeThresholdControlForTiffBinarization() throws Exception {
        //ExStart:ExposeThresholdControl
        //GistId:b20a0ec0e1ff0556aa20d12f486e1963
        Document doc = new Document(getMyDir() + "Rendering.docx");

        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.TIFF);
        saveOptions.setTiffCompression(TiffCompression.CCITT_3);
        saveOptions.setImageColorMode(ImageColorMode.GRAYSCALE);
        saveOptions.setTiffBinarizationMethod(ImageBinarizationMethod.FLOYD_STEINBERG_DITHERING);
        saveOptions.setThresholdForFloydSteinbergDithering((byte) 254);

        doc.save(getArtifactsDir() + "WorkingWithImageSaveOptions.ExposeThresholdControl.tiff", saveOptions);
        //ExEnd:ExposeThresholdControl
    }

    @Test
    public void getTiffPageRange() throws Exception {
        //ExStart:GetTiffPageRange
        //GistId:b20a0ec0e1ff0556aa20d12f486e1963
        Document doc = new Document(getMyDir() + "Rendering.docx");
        //ExStart:SaveAsTiff
        //GistId:b20a0ec0e1ff0556aa20d12f486e1963
        doc.save(getArtifactsDir() + "WorkingWithImageSaveOptions.MultipageTiff.tiff");
        //ExEnd:SaveAsTiff

        //ExStart:SaveAsTIFFUsingImageSaveOptions
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.TIFF);
        saveOptions.setPageSet(new PageSet(new PageRange(0, 1)));
        saveOptions.setTiffCompression(TiffCompression.CCITT_4);
        saveOptions.setResolution(160f);

        doc.save(getArtifactsDir() + "WorkingWithImageSaveOptions.GetTiffPageRange.tiff", saveOptions);
        //ExEnd:SaveAsTIFFUsingImageSaveOptions
        //ExEnd:GetTiffPageRange
    }

    @Test
    public void format1BppIndexed() throws Exception {
        //ExStart:Format1BppIndexed
        //GistId:83e5c469d0e72b5114fb8a05a1d01977
        Document doc = new Document(getMyDir() + "Rendering.docx");

        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.PNG);
        saveOptions.setPageSet(new PageSet(1));
        saveOptions.setImageColorMode(ImageColorMode.BLACK_AND_WHITE);
        saveOptions.setPixelFormat(ImagePixelFormat.FORMAT_1_BPP_INDEXED);

        doc.save(getArtifactsDir() + "WorkingWithImageSaveOptions.Format1BppIndexed.Png", saveOptions);
        //ExEnd:Format1BppIndexed
    }

    @Test
    public void getJpegPageRange() throws Exception {
        //ExStart:GetJpegPageRange
        //GistId:ebbb90d74ef57db456685052a18f8e86
        Document doc = new Document(getMyDir() + "Rendering.docx");

        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.JPEG);

        // Set the "PageSet" to "0" to convert only the first page of a document.
        options.setPageSet(new PageSet(0));

        // Change the image's brightness and contrast.
        // Both are on a 0-1 scale and are at 0.5 by default.
        options.setImageBrightness(0.3f);
        options.setImageContrast(0.7f);

        // Change the horizontal resolution.
        // The default value for these properties is 96.0, for a resolution of 96dpi.
        options.setHorizontalResolution(72f);

        doc.save(getArtifactsDir() + "WorkingWithImageSaveOptions.GetJpegPageRange.jpeg", options);
        //ExEnd:GetJpegPageRange
    }

    @Test
    //ExStart:PageSavingCallback
    public static void pageSavingCallback() throws Exception {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setPageSet(new PageSet(new PageRange(0, doc.getPageCount() - 1)));
        imageSaveOptions.setPageSavingCallback(new HandlePageSavingCallback());

        doc.save(getArtifactsDir() + "WorkingWithImageSaveOptions.PageSavingCallback.png", imageSaveOptions);
    }

    private static class HandlePageSavingCallback implements IPageSavingCallback {
        public void pageSaving(PageSavingArgs args) {
            args.setPageFileName(MessageFormat.format(getArtifactsDir() + "Page_{0}.png", args.getPageIndex()));
        }
    }
    //ExEnd:PageSavingCallback
}
