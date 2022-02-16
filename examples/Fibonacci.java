import java.util.Collection;

import nl.fountain.xelem.Address;
import nl.fountain.xelem.CellPointer;
import nl.fountain.xelem.XSerializer;
import nl.fountain.xelem.XelemException;
import nl.fountain.xelem.excel.Cell;
import nl.fountain.xelem.excel.Workbook;
import nl.fountain.xelem.excel.Worksheet;
import nl.fountain.xelem.excel.ss.SSCell;
import nl.fountain.xelem.excel.ss.XLWorkbook;

public class Fibonacci {

    public static void main(String[] args) {
        Workbook wb = createTheWorkbook();
        try {
            new XSerializer().serialize(wb);
            printWarnings(wb);
        } catch (XelemException e) {
            e.printStackTrace();
        }
    }
    
    private static Workbook createTheWorkbook() {
        Workbook wb = new XLWorkbook("Fibonacci");
        Worksheet sheet = wb.addSheet("ratio of Fibonacci numbers");
        sheet.addCell("Ratio of Fibonacci numbers", "title");
        
        // up to row 11 we want a white background
        for (int i = 1; i < 11; i++) {
            sheet.getRowAt(i).setStyleID("bg_white");
        }
        
        // add a heading and construct the formulas
        CellPointer cp = sheet.getCellPointer();
        cp.moveTo(10, 1);
        Address adrF1 = cp.getAddress();
        sheet.addCell("f1", "table_heading");
        
        Address adrF2 = cp.getAddress();
        sheet.addCell("f2", "table_heading");
        
        String formula1 = "=" + cp.getRefTo(adrF1) + "/" + cp.getRefTo(adrF2);
        sheet.addCell("ratio f1/f2", "table_heading");
        
        String formula2 = "=" + cp.getRefTo(adrF2) + "/" + cp.getRefTo(adrF1);
        sheet.addCell("ratio f2/f1", "table_heading");
        
        // put the relative formulas in cells
        Cell formulaCell1 = new SSCell();
        formulaCell1.setFormula(formula1);
        
        Cell formulaCell2 = new SSCell();
        formulaCell2.setFormula(formula2);
        
        // do the Fibonacci
        int f1 = 1;
        int f2 = 1;
        int f3;
        while (f1 < 1000000) {
            cp.moveCRLF();
            sheet.addCell(f1);
            sheet.addCell(f2);
            sheet.addCell(formulaCell1);
            sheet.addCell(formulaCell2);
            f3 = f1 + f2;
            f1 = f2;
            f2 = f3;
        }
        
        // the columns for the Fibonacci numbers should have a thousand seperator
        sheet.getColumnAt(1).setStyleID("dec_0");
        sheet.getColumnAt(2).setStyleID("dec_0");
        
        // the columns for the ratio should be wider
        sheet.getColumnAt(3).setWidth(75.0);
        sheet.getColumnAt(4).setWidth(75.0);
        
        // freeze the panes just below the table heading
        sheet.getWorksheetOptions().freezePanesAt(10, 0);
        
        return wb;
    }
    
    private static void printWarnings(Workbook wb) {
        Collection<String> warnings = wb.getWarnings();
        System.out.println("Created '" + wb.getFileName() 
                + "' with " + warnings.size() 
                + (warnings.size() == 1 ? " warning." : " warnings."));
        for (String s : warnings) {
            System.out.println(s);  
        }
    }
    
}
