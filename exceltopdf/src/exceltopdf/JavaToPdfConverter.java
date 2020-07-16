package exceltopdf;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.PageOrientation;
import com.grapecity.documents.excel.PaperSize;
import com.grapecity.documents.excel.SaveFileFormat;
import com.grapecity.documents.excel.Workbook;

public class JavaToPdfConverter {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Workbook workbook = new Workbook();
		
		workbook.open("data/Reports Samples_v.5.xlsx");
		
		IWorksheet sheet = workbook.getWorksheets().get(0);
		
		// Set bestFitColumns/bestFitRows as true.
//		sheet.getPageSetup().setBestFitColumns(true);
//		sheet.getPageSetup().setBestFitRows(true);
		
		// Set print gridline and heading.
//		sheet.getPageSetup().setPrintGridlines(true);
//		sheet.getPageSetup().setPrintHeadings(true);
		
		//Set page orientation.
		sheet.getPageSetup().setOrientation(PageOrientation.Landscape);
		
		//Set A4 paper size
		sheet.getPageSetup().setPaperSize(PaperSize.A4);
		
		//Set paper scaling
		//Method 1: Set percent scale 
//		sheet.getPageSetup().setIsPercentScale(true);
//		sheet.getPageSetup().setZoom(50);

		//Or Method 2: Fit to page's wide & tall
		sheet.getPageSetup().setIsPercentScale(false);
		sheet.getPageSetup().setFitToPagesWide(1);
		sheet.getPageSetup().setFitToPagesTall(1);

		sheet.save("data/Reports Samples_v.5.pdf", SaveFileFormat.Pdf);

	}
}
