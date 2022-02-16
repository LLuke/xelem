import java.util.Collection;

import nl.fountain.xelem.XSerializer;
import nl.fountain.xelem.XelemException;
import nl.fountain.xelem.excel.Workbook;
import nl.fountain.xelem.excel.ss.XLWorkbook;

public class HelloExcel {

    public static void main(String[] args) throws XelemException {
        Workbook wb = new XLWorkbook("HelloExcel");
        wb.addSheet().addCell("Hello Excel!");
        new XSerializer().serialize(wb);
        
        Collection<String> warnings = wb.getWarnings();
        System.out.println("Created '" + wb.getFileName() 
                + "' with " + warnings.size() 
                + (warnings.size() == 1 ? " warning." : " warnings."));
        for (String s : warnings) {
            System.out.println(s);            
        }
    }
}
