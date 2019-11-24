package com.project;

import com.itextpdf.kernel.events.Event;
import com.itextpdf.kernel.events.IEventHandler;
import com.itextpdf.kernel.events.PdfDocumentEvent;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfPage;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;
import com.itextpdf.layout.Canvas;
import com.itextpdf.layout.element.Image;

/**
 * This class happens to handle the image addition.
 * Note: The calling of this class manages the image as header or footer.
 * 
 * @author Sam Kar
 *
 */
public class ImageEventHandler implements IEventHandler {
	
    protected Image img = null;

    public ImageEventHandler(Image img) {
        this.img = img;
    }
    
	@SuppressWarnings("resource")
	public void handleEvent(Event event) {
		PdfDocumentEvent docEvent = (PdfDocumentEvent) event;
        PdfDocument pdfDoc = docEvent.getDocument();
        PdfPage page = docEvent.getPage();
        PdfCanvas aboveCanvas = new PdfCanvas(page.newContentStreamBefore(),
                page.getResources(), pdfDoc);
        Rectangle area = page.getPageSize();
        new Canvas(aboveCanvas, pdfDoc, area).add(img);
	}
}
