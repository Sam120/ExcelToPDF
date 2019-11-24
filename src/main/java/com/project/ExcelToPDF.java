package com.project;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.MalformedURLException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.io.font.constants.StandardFonts;
import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.events.PdfDocumentEvent;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.property.TextAlignment;

/**
 * Converts Excel into .pdf file with image header
 */
public class ExcelToPDF {
	public static Image img = null;
	public static Image img2 = null;
	
	/**
	 * Reads the left side image
	 * 
	 * @return
	 * @throws MalformedURLException
	 */
	public static Image getImage() throws MalformedURLException {
		ImageData imageData = ImageDataFactory.create("src/main/resources/Writing.png");
		img = new Image(imageData).scaleAbsolute(385, 81).setMarginTop(0);
		return img;
	}
	
	/**
	 * Reads the right side image
	 * 
	 * @return
	 * @throws MalformedURLException
	 */
	public static Image getImage2() throws MalformedURLException {
		ImageData imageData = ImageDataFactory.create("src/main/resources/Logo.png");
		img2 = new Image(imageData).scaleAbsolute(62, 81).setMarginTop(0).setMarginRight(0).setMarginLeft(1850);
		return img2;
	}
	
	@SuppressWarnings({ "unused", "resource" })
	public static void main(String[] args) throws IOException, MalformedURLException {
		
		FileInputStream inpDoc = new FileInputStream(new File("src/main/resources/SampleExcel.xlsx"));
		XSSFWorkbook wb = new XSSFWorkbook(inpDoc);
		int totalCellCount = 0;
		for (Cell cell : wb.getSheetAt(0).getRow(0)) {
			totalCellCount++;
		}
		
		PdfDocument pdfDoc = null;
		
		try {
			
			/**
			 * This creates the .pdf file in the scr/main/output folder of the project.
			 * To view the file, run the code and then refresh the project and copy the file into your desktop.
			 * It is better to change this path into your preferred folder.
			 */
			pdfDoc = new PdfDocument(new PdfWriter("src/main/output/OutputPDF.pdf"));
		}catch(Exception e) {
			System.out.println("The file is being used. Please close the file.");
		}
        
		/**
		 * Page size happens to be solving the case instead of using A4 and trying to rotate it.
		 * The page canvas width is wider enough to accommodate all columns. 
		 */
		//Document doc = new Document(pdfDoc, new PageSize(PageSize.A4).rotate());
        Document doc = new Document(pdfDoc, new PageSize(1920, 900));
        doc.setMargins(100, 1, 20, 1);
        
        /**
         * This denoted the position of the image to be added, here as header.
         */
        ImageEventHandler handler = new ImageEventHandler(getImage());
        pdfDoc.addEventHandler(PdfDocumentEvent.START_PAGE, handler);	//or END_PAGE as needed for footer
        ImageEventHandler handler2 = new ImageEventHandler(getImage2());
        pdfDoc.addEventHandler(PdfDocumentEvent.START_PAGE, handler2);  //or END_PAGE as needed for footer
        
        /**
         * Create table with auto aligning column width based on available page width
         */
        Table table = new Table(new float[totalCellCount]).useAllAvailableWidth();
        PdfFont bold = PdfFontFactory.createFont(StandardFonts.TIMES_BOLD);
        int i = 1;
        for(Sheet excelSheet : wb) {
        	for(Row excelRow : excelSheet) {
				for (Cell excelCell : excelRow) {
					if (i == 1) {
						com.itextpdf.layout.element.Cell tableCell = new com.itextpdf.layout.element.Cell();
						tableCell.add(new Paragraph(excelCell.getStringCellValue()))
								.setTextAlignment(TextAlignment.CENTER).setPadding(5).setFont(bold)
								.setBackgroundColor(new DeviceRgb(140, 221, 8));
						table.addCell(tableCell);
					} else {
						switch (excelCell.getCellType()) {
						case STRING:
							table.addCell(excelCell.getStringCellValue());
							break;
						case NUMERIC:
							table.addCell(String.valueOf(excelCell.getNumericCellValue()));
							break;
						default:
							break;
						}
					}
				}
        		i++;
        	}
        }
        doc.add(table);
        doc.close();
	}
}