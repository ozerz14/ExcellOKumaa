package exceldenokuma;

import java.io.FileInputStream;
import java.util.Iterator;
import java.util.Vector;
import java.io.File;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import org.apache.poi.ss.usermodel.DataFormatter;

public class ExceldenOkuma 
{
    public static void main(String[] args)
    {
        try
        {
            File file1 = new File("hatalar.txt");
			// if file doesnt exists, then create it
                if (!file1.exists())
                {
                    file1.createNewFile();
		}
 
		FileWriter fw = new FileWriter(file1.getAbsoluteFile());
		BufferedWriter bw = new BufferedWriter(fw);
                
                
                        
            FileInputStream file = new FileInputStream(new File("C:\\yeni.xls"));
 
            //Create Workbook instance holding reference to .xls file
            HSSFWorkbook workbook = new HSSFWorkbook(file);
 
            //Get first/desired sheet from the workbook
            HSSFSheet sheet = workbook.getSheetAt(0);
 
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                 int kontrol=0;
                while (cellIterator.hasNext())
                {
                    String s;
                    Double  d;
                    int x;
                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                    
                    switch (cell.getCellType())
                    {
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.println(new DataFormatter().formatCellValue(cell));
                            String strCellValue;
                            int i = (int)cell.getNumericCellValue();
                            strCellValue = String.valueOf(i); 
                            bw.write(strCellValue);
                            
                            kontrol++;
                           /* if(kontrol==1)
                            {
                            // bw.write(s);
                            bw.write("=");
                            }
                                   */
                            bw.write("   ");
                            break;
                            
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + "      ");
                            s=cell.getStringCellValue();
                            if(kontrol==1)
                            {
                            // bw.write(s);
                            bw.write("= ");
                            }
                            
                            if(kontrol!=1)
                            {
                            //System.out.print(cell.getStringCellValue() + "      ");
                            //s=cell.getStringCellValue();
                            bw.write(s);
                            bw.write("      ");
                            }
                            
                            
                            kontrol++;
                        break;
                    }
                }
                System.out.println("");
                bw.newLine();
               System.out.println("bir sonraki");
            }
            file.close();
            bw.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        
    }
}